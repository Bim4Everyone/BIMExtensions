# -*- coding: utf-8 -*-
"""Replaces current selection with elements inside the groups."""

import clr
clr.AddReference("dosymep.Revit.dll")

from Autodesk.Revit.DB import BuiltInParameter

import dosymep
clr.ImportExtensions(dosymep.Revit)

from pyrevit import revit, DB

doc = __revit__.ActiveUIDocument.Document

selection = revit.get_selection()


def get_group(element, group_elements=None):
    if not group_elements:
        group_elements = []

    for el in get_group_elements(element):
        if is_group(el):
            get_group(el, group_elements)
        else:
            group_elements.append(el)

    return group_elements


def get_group_elements(group):
    if is_group(group):
        for sub_element_id in group.GetMemberIds():
            sub_element = doc.GetElement(sub_element_id)

            phase_create_param = sub_element.GetParam(BuiltInParameter.PHASE_CREATED)
            if phase_create_param and not phase_create_param.IsReadOnly:
                yield sub_element


def is_group(element):
    return isinstance(element, DB.Group)


elements = []
for selected in selection:
    elements.extend(get_group(selected))

selection.set_to([e.Id for e in elements])
