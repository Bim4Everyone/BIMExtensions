# -*- coding: utf-8 -*-

import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *
from pyrevit import revit, DB
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document


def is_group(element):
    return isinstance(element, DB.Group)


def get_group(element, group_elements=None):
    if not group_elements:
        group_elements = []

    group_elements.append(element)
    for element in get_group_elements(element):
        if is_group(element):
            get_group(element, group_elements)

        group_elements.append(element)

    return group_elements


def get_group_elements(group):
    if is_group(group):
        for sub_element_id in group.GetMemberIds():
            yield doc.GetElement(sub_element_id)

        if not group.IsAttached:
            for sub_element_id in group.GetAvailableAttachedDetailGroupTypeIds():
                for group in doc.GetElement(sub_element_id).Groups:
                    yield group


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selection = revit.get_selection()

    elements = []
    for selected in selection:
        elements.extend(get_group(selected))

    selection.set_to(elements)


script_execute()
