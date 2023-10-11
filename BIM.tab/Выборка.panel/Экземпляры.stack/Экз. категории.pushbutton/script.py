# -*- coding: utf-8 -*-

from pyrevit import revit, DB
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    cl = DB.FilteredElementCollector(revit.doc, revit.active_view.Id) \
        .WhereElementIsNotElementType() \
        .ToElementIds()

    matchlist = []
    selCatList = set()

    selection = revit.get_selection()

    for el in selection:
        try:
            selCatList.add(el.Category.Name)
        except Exception:
            continue

    for elid in cl:
        el = revit.doc.GetElement(elid)
        try:
            # if el.ViewSpecific and ( el.Category.Name in selCatList):
            if el.Category.Name in selCatList:
                matchlist.append(elid)
        except Exception:
            continue

    selSet = []
    for elid in matchlist:
        selSet.append(elid)

    selection.set_to(selSet)
    revit.uidoc.RefreshActiveView()
    show_executed_script_notification()


script_execute()
