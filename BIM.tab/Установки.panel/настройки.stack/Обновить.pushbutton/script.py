# -*- coding: utf-8 -*-

import os
import pyrevit.coreutils.git as libgit
from pyrevit.versionmgr import updater
from pyrevit import EXEC_PARAMS
from pyrevit import script
from pyrevit import forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo

from dosymep_libs.bim4everyone import *

__cleanengine__ = True

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if EXEC_PARAMS.executed_from_ui:
        forms.alert('Вы уверены что хотите обновить?',
                          yes=True, no=True, exitscript=True)

    # пытаемся  обновится
    path = os.path.abspath(__file__)
    repo_path = libgit.libgit.Repository.Discover(path)
    repo_info = libgit.get_repo(repo_path)
    updater.update_repo(repo_info)

    logger = script.get_logger()
    results = script.get_results()

    # re-load pyrevit session.
    logger.info('Reloading....')
    sessionmgr.load_session()

    results.newsession = sessioninfo.get_session_uuid()


script_execute()