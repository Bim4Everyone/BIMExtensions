# -*- coding: utf-8 -*-

from pyrevit import EXEC_PARAMS
from pyrevit import script
from pyrevit import forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo

from dosymep_libs.bim4everyone import *

__cleanengine__ = True

@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    res = True
    if EXEC_PARAMS.executed_from_ui:
        res = forms.alert('Вы уверены что хотите обновить?',
                          yes=True, no=True)

    if res:
        logger = script.get_logger()
        results = script.get_results()

        # re-load pyrevit session.
        logger.info('Reloading....')
        sessionmgr.load_session()

        results.newsession = sessioninfo.get_session_uuid()
        show_executed_script_notification()


script_execute()