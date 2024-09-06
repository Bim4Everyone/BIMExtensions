# -*- coding: utf-8 -*-

from pyrevit import revit, DB
from pyrevit import framework
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selection = revit.get_selection()

    categories = [element.Category.Id for element in selection if element.Category]
    category_filter = DB.ElementMulticategoryFilter(framework.List[DB.ElementId](categories))

    elements = (DB.FilteredElementCollector(revit.doc, revit.active_view.Id)
                .WhereElementIsNotElementType()
                .WherePasses(category_filter)
                .ToElementIds())

    selection.set_to(elements)


script_execute()
