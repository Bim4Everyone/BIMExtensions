# -*- coding: utf-8 -*-

import clr
clr.AddReference('dosymep.Bim4Everyone.dll')

from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit import script
from pyrevit import forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo
from pyrevit.userconfig import user_config
from Autodesk.Revit.ApplicationServices import LanguageType

from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig


if HOST_APP.version == "2020":
    clr.AddReference('PlatformSettings.dll')
elif HOST_APP.version == "2021":
    clr.AddReference('PlatformSettings_2021.dll')
elif HOST_APP.version == "2022":
    clr.AddReference('PlatformSettings_2022.dll')

import PlatformSettings

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    user_config.auto_update = True
    user_config.check_updates = True
    user_config.save_changes()
    
    if __rvt__.Application.Language == LanguageType.Russian:
        user_config.user_locale = 'ru'
        
    if __rvt__.Application.Language == LanguageType.English_USA:
        user_config.user_locale = 'en_us'

    LoadPlatformSettings()


def LoadPlatformSettings():
    sharedParamsPath = None
    projectParamsPath = None
    if user_config.has_section("PlatformSettings"):
        try:
            sharedParamsPath = user_config.get_section("PlatformSettings").get_option("SharedParamsPath")
        except:
            pass
            
        try:
            projectParamsPath = user_config.get_section("PlatformSettings").get_option("ProjectParamsPath")
        except:
            pass
    
    SharedParamsConfig.LoadInstance(sharedParamsPath)
    ProjectParamsConfig.LoadInstance(projectParamsPath)


def OpenPlatrormSettings():
    settings = PlatformSettings.PlatformSettingsCommand()
    result = settings.Execute(EXEC_PARAMS.command_data)

    if result:
        user_config.reload()    
        LoadPlatformSettings()
        
        logger = script.get_logger()
        results = script.get_results()

        # re-load pyrevit session.
        logger.info('Reloading....')
        sessionmgr.load_session()

        results.newsession = sessioninfo.get_session_uuid()

if __name__ == '__main__':
    OpenPlatrormSettings()