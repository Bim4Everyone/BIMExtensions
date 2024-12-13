# -*- coding: utf-8 -*-

import os
import pyrevit.coreutils.git as libgit
from pyrevit.versionmgr import updater
from pyrevit import EXEC_PARAMS
from pyrevit import forms


from dosymep_libs.bim4everyone import *
import dosymep_libs

__cleanengine__ = True


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if EXEC_PARAMS.executed_from_ui:
        forms.alert('Вы уверены что хотите обновить?',
                          yes=True, no=True, exitscript=True)

    # пытаемся обновиться
    dosymep_libs.update_extension(__file__)


script_execute()
