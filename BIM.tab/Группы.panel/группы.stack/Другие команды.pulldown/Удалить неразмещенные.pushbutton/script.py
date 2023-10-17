# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *

from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    group_types = (FilteredElementCollector(revit.doc)
                   .OfClass(GroupType)
                   .ToElements())

    with revit.Transaction("BIM: Удаление не размещенных групп"):
        for group_type in group_types:  # type: GroupType
            if group_type.Groups.Size < 1:
                revit.doc.Delete(group_type.Id)


script_execute()
