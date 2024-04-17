# -*- coding: utf-8 -*-

import os
import pyrevit.coreutils.git as libgit
from pyrevit.versionmgr import updater
from pyrevit import EXEC_PARAMS
from pyrevit import forms


from dosymep_libs.bim4everyone import *

__cleanengine__ = True


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if EXEC_PARAMS.executed_from_ui:
        forms.alert('Вы уверены что хотите обновить?',
                          yes=True, no=True, exitscript=True)

    # пытаемся обновится
    path = os.path.abspath(__file__)
    repo_path = libgit.libgit.Repository.Discover(path)
    repo_info = libgit.get_repo(repo_path)
    updater.update_repo(repo_info)


script_execute()
