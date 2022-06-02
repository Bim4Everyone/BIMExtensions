# -*- coding: utf-8 -*-

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

    logger = script.get_logger()
    results = script.get_results()

    # re-load pyrevit session.
    logger.info('Reloading....')
    sessionmgr.load_session()

    results.newsession = sessioninfo.get_session_uuid()


script_execute()