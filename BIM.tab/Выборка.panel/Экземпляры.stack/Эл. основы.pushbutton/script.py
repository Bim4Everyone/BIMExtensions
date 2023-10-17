# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selection = revit.get_selection()
    selected_walls = {wall.Id for wall in selection.elements if isinstance(wall, Wall)}
    if not selected_walls:
        with forms.WarningBar(title="Выберите стены"):
            picked_walls = revit.pick_elements_by_category(BuiltInCategory.OST_Walls)
            selected_walls = {wall.Id for wall in picked_walls}

    wall_sweeps = FilteredElementCollector(revit.doc).OfClass(WallSweep).ToElements()
    wall_sweeps = [sweep.Id for sweep in wall_sweeps
                   if selected_walls.intersection(sweep.GetHostIds())]

    family_instances = FilteredElementCollector(revit.doc).OfClass(FamilyInstance).ToElements()
    family_instances = [instance.Id for instance in family_instances
                        if instance.Host and instance.Host.Id in selected_walls]

    selection.set_to(selected_walls.union(wall_sweeps, family_instances))


script_execute()
