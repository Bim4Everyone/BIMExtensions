# -*- coding: utf-8 -*-

import dosymep_libs
dosymep_libs.load_assemblies()

import clr
clr.AddReference('dosymep.Revit.dll')
clr.AddReference('dosymep.Bim4Everyone.dll')

from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit.userconfig import user_config
from Autodesk.Revit.ApplicationServices import LanguageType

from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig

if HOST_APP.version == "2020":
    clr.AddReference("PlatformSettings.dll")
else:
    clr.AddReference("PlatformSettings_{}.dll".format(HOST_APP.version))

import PlatformSettings


def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    user_config.auto_update = True
    user_config.check_updates = True
    user_config.save_changes()
    
    if __rvt__.Application.Language == LanguageType.Russian:
        user_config.user_locale = 'ru'
        
    if __rvt__.Application.Language == LanguageType.English_USA:
        user_config.user_locale = 'en_us'

    load_platform_settings()

def get_config_path(section, option):
    try:
        if user_config.has_section(section):
            return user_config.get_section(section).get_option(option)
    except:
        pass


def load_platform_settings():
    shared_params_path = get_config_path("PlatformSettings", "SharedParamsPath")
    project_params_path = get_config_path("PlatformSettings", "ProjectParamsPath")

    SharedParamsConfig.LoadInstance(shared_params_path)
    ProjectParamsConfig.LoadInstance(project_params_path)

def open_platform_settings():
    settings = PlatformSettings.PlatformSettingsCommand()
    result = settings.Execute(EXEC_PARAMS.command_data)

    if result:
        user_config.reload()
        load_platform_settings()

        from pyrevit.loader.sessionmgr import execute_command
        execute_command("01dotbim-bim-установки-settings-reload")

if __name__ == '__main__':
    open_platform_settings()