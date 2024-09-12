# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import revit
from pyrevit import framework
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selection = revit.get_selection()
    group_types = [element for element in selection.elements if isinstance(element, GroupType)]

    if not group_types:
        forms.alert("Должен быть выбран хотя бы один типоразмер группы.", exitscript=True)

    # group_types.SelectMany(item => item.Groups)
    groups = [g.Id for group_type in group_types for g in group_type.Groups]

    if not groups:
        forms.alert("Не были найдены экземпляры групп в типоразмерах.", exitscript=True)

    revit.uidoc.ShowElements(framework.List[ElementId](groups))
    selection.set_to(groups)


script_execute()
