/*
 * obs-powerpoint-slides
 * Fuente de OBS que muestra presentaciones (.pptx/.ppt/.odp) diapositiva a diapositiva.
 *
 * Funcionamiento:
 *  - Un hilo en segundo plano convierte la presentación a un PNG por diapositiva.
 *      Windows : PowerPoint (COM) mediante PowerShell
 *      Linux/macOS : LibreOffice (soffice) + pdftoppm (poppler)
 *  - La fuente carga el PNG de la diapositiva actual como textura y lo dibuja.
 *  - Botones en las propiedades y atajos de teclado para avanzar/retroceder.
 */

#include <obs-module.h>
#include <graphics/image-file.h>
#include <util/platform.h>
#include <util/dstr.h>
#include <util/bmem.h>
#include <pthread.h>
#include <stdlib.h>
#include <string.h>
#include <ctype.h>

#ifdef _WIN32
#include <windows.h>
#endif

struct ppt_source {
	obs_source_t *source;
	pthread_mutex_t mutex;
	pthread_t thread;
	bool thread_valid;
	bool running;
	bool restart;

	char *file;
	char *cache_dir;
	char *status;
	int width;
	int height;

	char **slides;
	size_t slide_count;
	bool list_dirty;

	int slide; /* base 1 */
	bool slide_dirty;

	gs_image_file_t image;
	bool image_loaded;

	obs_hotkey_id hk_next;
	obs_hotkey_id hk_prev;
};

/* ------------------------------------------------------------------ */
/* Utilidades                                                          */
/* ------------------------------------------------------------------ */

static void set_status_locked(struct ppt_source *s, const char *text)
{
	bfree(s->status);
	s->status = bstrdup(text);
}

static char *cache_dir_for(const char *file)
{
	unsigned int h = 5381;
	for (const unsigned char *p = (const unsigned char *)file; *p; p++)
		h = ((h << 5) + h) + *p;

	struct dstr rel = {0};
	dstr_printf(&rel, "powerpoint-cache/%08x", h);
	char *dir = obs_module_config_path(rel.array);
	dstr_free(&rel);
	if (dir)
		os_mkdirs(dir);
	return dir;
}

static void clear_cache(const char *dir)
{
	os_dir_t *d = os_opendir(dir);
	if (!d)
		return;
	struct os_dirent *e;
	while ((e = os_readdir(d))) {
		if (e->directory)
			continue;
		const char *ext = strrchr(e->d_name, '.');
		if (ext && (astrcmpi(ext, ".png") == 0 || astrcmpi(ext, ".pdf") == 0 ||
			    astrcmpi(ext, ".txt") == 0)) {
			struct dstr p = {0};
			dstr_printf(&p, "%s/%s", dir, e->d_name);
			os_unlink(p.array);
			dstr_free(&p);
		}
	}
	os_closedir(d);
}

/* Ejecuta un comando sin mostrar ventana y espera a que termine. */
static bool run_hidden(const char *cmdline)
{
#ifdef _WIN32
	wchar_t *w = NULL;
	if (os_utf8_to_wcs_ptr(cmdline, 0, &w) == 0)
		return false;

	STARTUPINFOW si;
	PROCESS_INFORMATION pi;
	memset(&si, 0, sizeof(si));
	si.cb = sizeof(si);
	memset(&pi, 0, sizeof(pi));

	BOOL ok = CreateProcessW(NULL, w, NULL, NULL, FALSE, CREATE_NO_WINDOW, NULL, NULL, &si, &pi);
	bfree(w);
	if (!ok) {
		blog(LOG_WARNING, "[powerpoint] no se pudo iniciar PowerShell (error %lu)", GetLastError());
		return false;
	}

	WaitForSingleObject(pi.hProcess, 5 * 60 * 1000);
	DWORD code = 1;
	GetExitCodeProcess(pi.hProcess, &code);
	CloseHandle(pi.hProcess);
	CloseHandle(pi.hThread);
	return code == 0;
#else
	return system(cmdline) == 0;
#endif
}

static bool convert_presentation(const char *file, const char *dir, int width, int height, char **err_msg)
{
	if (err_msg)
		*err_msg = NULL;
	struct dstr script = {0};
	struct dstr cmd = {0};
	bool ok = false;

	clear_cache(dir);

#ifdef _WIN32
	static const char ps[] =
		"param([string]$In,[string]$Out,[int]$W,[int]$H)\r\n"
		"$ErrorActionPreference='Stop'\r\n"
		"$app = $null\r\n"
		"try {\r\n"
		"  $In = [System.IO.Path]::GetFullPath($In)\r\n"
		"  $Out = [System.IO.Path]::GetFullPath($Out)\r\n"
		"  if (-not (Test-Path 'Registry::HKEY_CLASSES_ROOT\\PowerPoint.Application')) {\r\n"
		"    throw 'Microsoft PowerPoint (version de escritorio) no esta instalado en este equipo.'\r\n"
		"  }\r\n"
		"  $copy = Join-Path $Out ('source' + [System.IO.Path]::GetExtension($In))\r\n"
		"  Copy-Item -LiteralPath $In -Destination $copy -Force\r\n"
		"  Unblock-File -LiteralPath $copy\r\n"
		"  $app = New-Object -ComObject PowerPoint.Application\r\n"
		"  try { $pres = $app.Presentations.Open($copy, -1, 0, 0) }\r\n"
		"  catch { $pres = $app.Presentations.Open($copy, -1, 0, -1) }\r\n"
		"  $pres.Export($Out, 'PNG', $W, $H)\r\n"
		"  $pres.Close()\r\n"
		"} catch {\r\n"
		"  $err = Join-Path ([System.IO.Path]::GetFullPath($Out)) 'error.txt'\r\n"
		"  ($_.Exception.Message) | Out-File -LiteralPath $err -Encoding utf8\r\n"
		"  exit 1\r\n"
		"} finally {\r\n"
		"  if ($app -ne $null -and $app.Presentations.Count -eq 0) { $app.Quit() }\r\n"
		"}\r\n";

	dstr_printf(&script, "%s/convert.ps1", dir);
	os_quick_write_utf8_file(script.array, ps, strlen(ps), false);
	dstr_printf(&cmd,
		    "powershell.exe -NoProfile -ExecutionPolicy Bypass -File \"%s\" \"%s\" \"%s\" %d %d",
		    script.array, file, dir, width, height);
#else
	static const char sh[] =
		"#!/bin/sh\n"
		"IN=\"$1\"; OUT=\"$2\"; DPI=\"$3\"\n"
		"PATH=\"$PATH:/opt/homebrew/bin:/usr/local/bin\"\n"
		"SOFFICE=$(command -v soffice || command -v libreoffice || echo /Applications/LibreOffice.app/Contents/MacOS/soffice)\n"
		"\"$SOFFICE\" --headless --convert-to pdf --outdir \"$OUT\" \"$IN\" || exit 1\n"
		"BASE=$(basename \"$IN\"); PDF=\"$OUT/${BASE%.*}.pdf\"\n"
		"pdftoppm -png -r \"$DPI\" \"$PDF\" \"$OUT/slide\" || exit 1\n";

	dstr_printf(&script, "%s/convert.sh", dir);
	os_quick_write_utf8_file(script.array, sh, strlen(sh), false);
	dstr_printf(&cmd, "/bin/sh \"%s\" \"%s\" \"%s\" %d", script.array, file, dir, (width * 3) / 40);
#endif

	blog(LOG_INFO, "[powerpoint] ejecutando: %s", cmd.array);
	ok = run_hidden(cmd.array);
	if (!ok) {
		struct dstr ep = {0};
		dstr_printf(&ep, "%s/error.txt", dir);
		char *msg = os_quick_read_utf8_file(ep.array);
		blog(LOG_WARNING, "[powerpoint] la conversión falló: %s",
		     msg ? msg : "(sin detalles; revisa que PowerPoint esté instalado)");
		if (err_msg)
			*err_msg = msg;
		else
			bfree(msg);
		dstr_free(&ep);
	}
	dstr_free(&script);
	dstr_free(&cmd);
	return ok;
}

/* ------------------------------------------------------------------ */
/* Lectura de diapositivas generadas                                   */
/* ------------------------------------------------------------------ */

struct slide_entry {
	int n;
	char *path;
};

static int slide_number(const char *name)
{
	const char *dot = strrchr(name, '.');
	const char *p = dot ? dot : name + strlen(name);
	while (p > name && !isdigit((unsigned char)p[-1]))
		p--;
	while (p > name && isdigit((unsigned char)p[-1]))
		p--;
	return atoi(p);
}

static int cmp_entry(const void *a, const void *b)
{
	return ((const struct slide_entry *)a)->n - ((const struct slide_entry *)b)->n;
}

static size_t scan_slides(const char *dir, char ***out)
{
	struct slide_entry *list = NULL;
	size_t count = 0;

	os_dir_t *d = os_opendir(dir);
	if (!d)
		return 0;

	struct os_dirent *e;
	while ((e = os_readdir(d))) {
		if (e->directory)
			continue;
		const char *ext = strrchr(e->d_name, '.');
		if (!ext || astrcmpi(ext, ".png") != 0)
			continue;

		list = brealloc(list, (count + 1) * sizeof(*list));
		struct dstr p = {0};
		dstr_printf(&p, "%s/%s", dir, e->d_name);
		list[count].n = slide_number(e->d_name);
		list[count].path = p.array;
		count++;
	}
	os_closedir(d);

	if (count == 0) {
		*out = NULL;
		return 0;
	}

	qsort(list, count, sizeof(*list), cmp_entry);

	char **paths = bzalloc(count * sizeof(char *));
	for (size_t i = 0; i < count; i++)
		paths[i] = list[i].path;
	bfree(list);

	*out = paths;
	return count;
}

static void free_slide_list(char **slides, size_t count)
{
	for (size_t i = 0; i < count; i++)
		bfree(slides[i]);
	bfree(slides);
}

/* ------------------------------------------------------------------ */
/* Hilo de conversión                                                  */
/* ------------------------------------------------------------------ */

static void *convert_thread(void *param)
{
	struct ppt_source *s = param;

	for (;;) {
		pthread_mutex_lock(&s->mutex);
		char *file = bstrdup(s->file);
		char *dir = bstrdup(s->cache_dir);
		int w = s->width;
		int h = s->height;
		s->restart = false;
		set_status_locked(s, "Convirtiendo presentación…");
		pthread_mutex_unlock(&s->mutex);

		char *err = NULL;
		bool ok = file && dir && convert_presentation(file, dir, w, h, &err);
		char **list = NULL;
		size_t n = 0;
		if (ok)
			n = scan_slides(dir, &list);
		blog(LOG_INFO, "[powerpoint] conversión %s, %d diapositivas encontradas en %s",
		     ok ? "correcta" : "fallida", (int)n, dir ? dir : "(sin carpeta)");

		bfree(file);
		bfree(dir);

		pthread_mutex_lock(&s->mutex);
		if (s->restart) {
			/* Cambiaron los ajustes mientras convertía: repetir. */
			pthread_mutex_unlock(&s->mutex);
			free_slide_list(list, n);
			bfree(err);
			continue;
		}

		if (s->slides)
			free_slide_list(s->slides, s->slide_count);
		s->slides = list;
		s->slide_count = n;
		s->list_dirty = true;

		if (!ok) {
			struct dstr msg = {0};
			dstr_printf(&msg, "Error al convertir: %s",
				    err ? err : "sin detalles (revisa el registro de OBS y busca [powerpoint])");
			set_status_locked(s, msg.array);
			dstr_free(&msg);
		}
		else if (n == 0)
			set_status_locked(s, "No se generaron diapositivas.");
		else {
			struct dstr msg = {0};
			dstr_printf(&msg, "%d diapositivas cargadas.", (int)n);
			set_status_locked(s, msg.array);
			dstr_free(&msg);
		}

		bfree(err);
		s->running = false;
		pthread_mutex_unlock(&s->mutex);
		break;
	}
	return NULL;
}

/* Llamar con el mutex tomado. */
static void request_convert_locked(struct ppt_source *s)
{
	if (!s->file || !*s->file)
		return;

	if (s->running) {
		s->restart = true;
		return;
	}

	if (s->thread_valid) {
		pthread_join(s->thread, NULL);
		s->thread_valid = false;
	}

	s->running = true;
	s->restart = false;
	set_status_locked(s, "Convirtiendo presentación…");
	s->thread_valid = pthread_create(&s->thread, NULL, convert_thread, s) == 0;
	if (!s->thread_valid)
		s->running = false;
}

/* ------------------------------------------------------------------ */
/* Navegación                                                          */
/* ------------------------------------------------------------------ */

static void go_to_slide(struct ppt_source *s, int slide, bool relative)
{
	pthread_mutex_lock(&s->mutex);
	if (s->slide_count == 0) {
		pthread_mutex_unlock(&s->mutex);
		return;
	}
	int target = relative ? s->slide + slide : slide;
	if (target < 1)
		target = 1;
	if (target > (int)s->slide_count)
		target = (int)s->slide_count;
	s->slide = target;
	s->slide_dirty = true;
	pthread_mutex_unlock(&s->mutex);

	obs_data_t *settings = obs_source_get_settings(s->source);
	obs_data_set_int(settings, "slide", target);
	obs_data_release(settings);
}

static void hotkey_next(void *data, obs_hotkey_id id, obs_hotkey_t *hotkey, bool pressed)
{
	UNUSED_PARAMETER(id);
	UNUSED_PARAMETER(hotkey);
	if (pressed)
		go_to_slide(data, 1, true);
}

static void hotkey_prev(void *data, obs_hotkey_id id, obs_hotkey_t *hotkey, bool pressed)
{
	UNUSED_PARAMETER(id);
	UNUSED_PARAMETER(hotkey);
	if (pressed)
		go_to_slide(data, -1, true);
}

/* ------------------------------------------------------------------ */
/* Callbacks de la fuente                                              */
/* ------------------------------------------------------------------ */

static const char *ppt_get_name(void *unused)
{
	UNUSED_PARAMETER(unused);
	return "Presentación de PowerPoint";
}

static void ppt_update(void *data, obs_data_t *settings)
{
	struct ppt_source *s = data;
	const char *file = obs_data_get_string(settings, "file");
	int w = (int)obs_data_get_int(settings, "width");
	int h = (int)obs_data_get_int(settings, "height");
	int slide = (int)obs_data_get_int(settings, "slide");

	pthread_mutex_lock(&s->mutex);

	bool changed = !s->file || strcmp(s->file, file) != 0 || w != s->width || h != s->height;
	if (changed) {
		bfree(s->file);
		bfree(s->cache_dir);
		s->file = bstrdup(file);
		s->cache_dir = *file ? cache_dir_for(file) : NULL;
		s->width = w;
		s->height = h;
		request_convert_locked(s);
	}

	if (slide != s->slide) {
		s->slide = slide;
		s->slide_dirty = true;
	}

	pthread_mutex_unlock(&s->mutex);
}

static void *ppt_create(obs_data_t *settings, obs_source_t *source)
{
	struct ppt_source *s = bzalloc(sizeof(*s));
	s->source = source;
	s->slide = 1;
	pthread_mutex_init(&s->mutex, NULL);
	s->status = bstrdup("Selecciona una presentación.");

	s->hk_next = obs_hotkey_register_source(source, "obs_powerpoint.next",
						"PowerPoint: diapositiva siguiente", hotkey_next, s);
	s->hk_prev = obs_hotkey_register_source(source, "obs_powerpoint.prev",
						"PowerPoint: diapositiva anterior", hotkey_prev, s);

	ppt_update(s, settings);
	return s;
}

static void ppt_destroy(void *data)
{
	struct ppt_source *s = data;

	if (s->thread_valid)
		pthread_join(s->thread, NULL);

	obs_enter_graphics();
	if (s->image_loaded)
		gs_image_file_free(&s->image);
	obs_leave_graphics();

	if (s->slides)
		free_slide_list(s->slides, s->slide_count);
	bfree(s->file);
	bfree(s->cache_dir);
	bfree(s->status);
	pthread_mutex_destroy(&s->mutex);
	bfree(s);
}

static void ppt_defaults(obs_data_t *settings)
{
	obs_data_set_default_int(settings, "width", 1920);
	obs_data_set_default_int(settings, "height", 1080);
	obs_data_set_default_int(settings, "slide", 1);
}

static bool btn_prev(obs_properties_t *props, obs_property_t *prop, void *data)
{
	UNUSED_PARAMETER(props);
	UNUSED_PARAMETER(prop);
	go_to_slide(data, -1, true);
	return true;
}

static bool btn_next(obs_properties_t *props, obs_property_t *prop, void *data)
{
	UNUSED_PARAMETER(props);
	UNUSED_PARAMETER(prop);
	go_to_slide(data, 1, true);
	return true;
}

static void refresh_info(struct ppt_source *s, obs_properties_t *props)
{
	obs_property_t *info = obs_properties_get(props, "info");
	if (!info)
		return;
	pthread_mutex_lock(&s->mutex);
	obs_property_set_description(info, s->status ? s->status : "");
	pthread_mutex_unlock(&s->mutex);
}

static bool btn_reload(obs_properties_t *props, obs_property_t *prop, void *data)
{
	UNUSED_PARAMETER(prop);
	struct ppt_source *s = data;
	pthread_mutex_lock(&s->mutex);
	if (s->file && *s->file && !s->running)
		set_status_locked(s, "Convirtiendo presentación…");
	request_convert_locked(s);
	pthread_mutex_unlock(&s->mutex);
	refresh_info(s, props);
	return true;
}

static bool btn_status(obs_properties_t *props, obs_property_t *prop, void *data)
{
	UNUSED_PARAMETER(prop);
	refresh_info(data, props);
	return true;
}

static obs_properties_t *ppt_properties(void *data)
{
	struct ppt_source *s = data;
	obs_properties_t *props = obs_properties_create();

	obs_properties_add_path(props, "file", "Presentación", OBS_PATH_FILE,
				"Presentaciones (*.pptx *.ppt *.pptm *.odp);;Todos los archivos (*.*)", NULL);

	obs_property_t *info = obs_properties_add_text(props, "info", "Estado", OBS_TEXT_INFO);
	if (s) {
		pthread_mutex_lock(&s->mutex);
		obs_property_set_description(info, s->status ? s->status : "");
		pthread_mutex_unlock(&s->mutex);
	}
	obs_properties_add_button(props, "status", "Actualizar estado", btn_status);

	obs_properties_add_int(props, "slide", "Diapositiva", 1, 9999, 1);
	obs_properties_add_button(props, "prev", "◀ Anterior", btn_prev);
	obs_properties_add_button(props, "next", "Siguiente ▶", btn_next);
	obs_properties_add_button(props, "reload", "Recargar presentación", btn_reload);

	obs_properties_add_int(props, "width", "Ancho de render (px)", 320, 7680, 1);
	obs_properties_add_int(props, "height", "Alto de render (px)", 240, 4320, 1);
	return props;
}

static void load_slide_image(struct ppt_source *s, const char *path)
{
	obs_enter_graphics();
	if (s->image_loaded) {
		gs_image_file_free(&s->image);
		s->image_loaded = false;
	}
	if (path) {
		gs_image_file_init(&s->image, path);
		gs_image_file_init_texture(&s->image);
		s->image_loaded = s->image.loaded;
	}
	obs_leave_graphics();
}

static void ppt_tick(void *data, float seconds)
{
	UNUSED_PARAMETER(seconds);
	struct ppt_source *s = data;
	char *path = NULL;
	bool update_image = false;

	pthread_mutex_lock(&s->mutex);
	if (s->list_dirty) {
		s->list_dirty = false;
		s->slide_dirty = true;
	}
	if (s->slide_dirty) {
		s->slide_dirty = false;
		update_image = true;
		if (s->slide_count > 0) {
			if (s->slide < 1)
				s->slide = 1;
			if (s->slide > (int)s->slide_count)
				s->slide = (int)s->slide_count;
			path = bstrdup(s->slides[s->slide - 1]);
		}
	}
	pthread_mutex_unlock(&s->mutex);

	if (update_image) {
		load_slide_image(s, path);
		bfree(path);
	}
}

static void ppt_render(void *data, gs_effect_t *unused)
{
	UNUSED_PARAMETER(unused);
	struct ppt_source *s = data;
	if (!s->image_loaded || !s->image.texture)
		return;

	gs_effect_t *effect = obs_get_base_effect(OBS_EFFECT_DEFAULT);
	gs_texture_t *tex = s->image.texture;

	const bool previous = gs_framebuffer_srgb_enabled();
	gs_enable_framebuffer_srgb(true);
	gs_blend_state_push();
	gs_blend_function(GS_BLEND_ONE, GS_BLEND_INVSRCALPHA);

	gs_eparam_t *param = gs_effect_get_param_by_name(effect, "image");
	gs_effect_set_texture_srgb(param, tex);
	while (gs_effect_loop(effect, "Draw"))
		gs_draw_sprite(tex, 0, s->image.cx, s->image.cy);

	gs_blend_state_pop();
	gs_enable_framebuffer_srgb(previous);
}

static uint32_t ppt_width(void *data)
{
	struct ppt_source *s = data;
	return s->image_loaded ? s->image.cx : (uint32_t)s->width;
}

static uint32_t ppt_height(void *data)
{
	struct ppt_source *s = data;
	return s->image_loaded ? s->image.cy : (uint32_t)s->height;
}

struct obs_source_info powerpoint_source_info = {
	.id = "powerpoint_slides",
	.type = OBS_SOURCE_TYPE_INPUT,
	.output_flags = OBS_SOURCE_VIDEO,
	.get_name = ppt_get_name,
	.create = ppt_create,
	.destroy = ppt_destroy,
	.update = ppt_update,
	.get_defaults = ppt_defaults,
	.get_properties = ppt_properties,
	.video_tick = ppt_tick,
	.video_render = ppt_render,
	.get_width = ppt_width,
	.get_height = ppt_height,
};
