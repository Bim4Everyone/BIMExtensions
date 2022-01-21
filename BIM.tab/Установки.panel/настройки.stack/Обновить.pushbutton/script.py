# -*- coding: utf-8 -*-
"""Reload pyRevit into new session."""

from pyrevit import EXEC_PARAMS
from pyrevit import script
from pyrevit import forms
from pyrevit.loader import sessionmgr
from pyrevit.loader import sessioninfo

__title__ = "."
__cleanengine__ = True
__doc__ = 'Searches the script folders and create buttons ' \
          'for the new script or newly installed extensions.'


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
