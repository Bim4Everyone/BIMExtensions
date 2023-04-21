# -*- coding: utf-8 -*-

import clr
clr.AddReference('dosymep.Revit.dll')
clr.AddReference('dosymep.Bim4Everyone.dll')

from System import InvalidOperationException, OperationCanceledException

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.ApplicationServices import LanguageType

from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit.userconfig import user_config

from dosymep.Revit import *
from dosymep.Bim4Everyone import *

from dosymep_libs.bim4everyone import *

def __selfinit__(script_cmp, ui_button_cmp, __rvt__):
    user_config.auto_update = True
    user_config.check_updates = True
    user_config.save_changes()

    if HOST_APP.language == LanguageType.Russian:
        user_config.user_locale = 'ru'

    if HOST_APP.language == LanguageType.English_USA:
        user_config.user_locale = 'en_us'


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    try:
        invoke_command(PlatformCommandIds.PlatformSettingsCommandId)

        user_config.reload()
        from pyrevit.loader.sessionmgr import execute_command
        execute_command("01dotbim-bim-установки-настройки-обновить")
    except OperationCanceledException:
        show_canceled_script_notification()


if __name__ == '__main__':
    script_execute()
