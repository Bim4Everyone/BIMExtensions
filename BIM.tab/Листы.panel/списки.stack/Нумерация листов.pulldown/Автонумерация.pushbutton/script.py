# -*- coding: utf-8 -*-

from pySpeech.ViewSheets import *

from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    order_view = OrderViewSheetModel(DocumentRepository(__revit__))

    order_view.LoadViewSheets()
    order_view.CheckUniquesNames()
    order_view.OrderViewSheets()


script_execute()