# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *

from pyrevit.framework import *
from pyrevit import forms
from pyrevit import script
from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    error_collector = {}

    selection = __revit__.ActiveUIDocument.Selection.GetElementIds()
    temp = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()
    wallsweeps = FilteredElementCollector(doc).OfClass(WallSweep).ToElements()

    associate = []
    walls = []
    for elid in selection:
        el = doc.GetElement(elid)
        if isinstance(el, Wall):
            walls.append(elid)
            associate.append(elid)

    for t in temp:
        if t.Host:
            for wall in walls:
                if str(t.Host.Id) == wall.ToString():
                    associate.append(t.Id)
                    break

    for ws in wallsweeps:
        for wall in walls:
            if wall in ws.GetHostIds():
                associate.append(ws.Id)
                break

    selection = revit.get_selection()
    selection.set_to(associate)
    show_executed_script_notification()


script_execute()
