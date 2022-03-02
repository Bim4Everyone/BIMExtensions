# -*- coding: utf-8 -*-

import clr
clr.AddReference('dosymep.Revit.dll')
clr.AddReference('dosymep.Bim4Everyone.dll')

from Autodesk.Revit.ApplicationServices import LanguageType

from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit.userconfig import user_config

if HOST_APP.version == "2020":
    clr.AddReference("PlatformSettings.dll")
else:
    clr.AddReference("PlatformSettings_{}.dll".format(HOST_APP.version))

import PlatformSettings

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    if HOST_APP.language == LanguageType.Russian:
        user_config.user_locale = 'ru'

    if HOST_APP.language == LanguageType.English_USA:
        user_config.user_locale = 'en_us'


def open_platform_settings():
    settings = PlatformSettings.PlatformSettingsCommand()
    result = settings.Execute(EXEC_PARAMS.command_data)

    if result:
        user_config.reload()

        from pyrevit.loader.sessionmgr import execute_command
        execute_command("01dotbim-bim-установки-настройки-обновить")

if __name__ == '__main__':
    open_platform_settings()
