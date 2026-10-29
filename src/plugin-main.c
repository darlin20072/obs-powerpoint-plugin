#include <obs-module.h>

OBS_DECLARE_MODULE()

extern struct obs_source_info powerpoint_source_info;

const char *obs_module_description(void)
{
	return "Fuente para mostrar presentaciones de PowerPoint en OBS";
}

bool obs_module_load(void)
{
	obs_register_source(&powerpoint_source_info);
	return true;
}
