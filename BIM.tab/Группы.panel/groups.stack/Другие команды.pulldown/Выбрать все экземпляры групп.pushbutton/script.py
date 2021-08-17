# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
import clr
clr.AddReference('System')
from System.Collections.Generic import List 

from Autodesk.Revit.DB import  GroupType, FilteredElementCollector, Transaction, TransactionGroup, ElementId
__doc__ = 'Выделяет все элементы, относящиеся к выбранным группам'
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument



groupTypes = [ doc.GetElement(elId) for elId in __revit__.ActiveUIDocument.Selection.GetElementIds() if isinstance(doc.GetElement( elId ), GroupType)]

groups = []

for groupType in groupTypes:

	groups = groups + [x.Id for x in groupType.Groups]


groupsList = List[ElementId](groups)


__revit__.ActiveUIDocument.Selection.SetElementIds(groupsList)