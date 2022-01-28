# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *

from pyrevit import revit

doc = __revit__.ActiveUIDocument.Document
group_types = [x for x in FilteredElementCollector(doc).OfClass(GroupType).ToElements()]

with revit.Transaction("BIM: Удаление не размещенных групп"):
    for group_type in group_types:  # type: GroupType
        if group_type.Groups.Size < 1:
            doc.Delete(group_type.Id)
