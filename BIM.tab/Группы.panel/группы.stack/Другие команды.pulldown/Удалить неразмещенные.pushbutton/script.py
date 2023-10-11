# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *

from pyrevit import revit
from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    doc = __revit__.ActiveUIDocument.Document
    group_types = [x for x in FilteredElementCollector(doc).OfClass(GroupType).ToElements()]

    with revit.Transaction("BIM: Удаление не размещенных групп"):
        for group_type in group_types:  # type: GroupType
            if group_type.Groups.Size < 1:
                doc.Delete(group_type.Id)

    show_executed_script_notification()


script_execute()
