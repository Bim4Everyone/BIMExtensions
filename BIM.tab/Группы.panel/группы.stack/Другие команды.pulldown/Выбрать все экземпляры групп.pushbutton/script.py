# -*- coding: utf-8 -*-

import clr
clr.AddReference('System')

from System.Collections.Generic import List

from pyrevit import forms
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

group_types = [doc.GetElement(elId) for elId in uidoc.Selection.GetElementIds()
               if isinstance(doc.GetElement(elId), GroupType)]

if not group_types:
    forms.alert("Должен быть выбран хотя бы один типоразмер группы.", exitscript=True)

groups = [g.Id for group_type in group_types
          for g in group_type.Groups]

if not groups:
    forms.alert("Не были найдены экземпляры групп в типоразмерах.", exitscript=True)

__revit__.ActiveUIDocument.ShowElements(List[ElementId](groups))
__revit__.ActiveUIDocument.Selection.SetElementIds(List[ElementId](groups))
