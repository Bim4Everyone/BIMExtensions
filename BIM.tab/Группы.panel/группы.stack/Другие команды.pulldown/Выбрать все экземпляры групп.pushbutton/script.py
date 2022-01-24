# -*- coding: utf-8 -*-

import clr
clr.AddReference('System')

from System.Collections.Generic import List

from Autodesk.Revit.DB import *


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

group_types = [doc.GetElement(elId) for elId in __revit__.ActiveUIDocument.Selection.GetElementIds()
			   if isinstance(doc.GetElement(elId), GroupType)]


groups = [ g.Id for group_type in group_types
		  for g in group_type.Groups ]

__revit__.ActiveUIDocument.ShowElements(List[ElementId](groups))
__revit__.ActiveUIDocument.Selection.SetElementIds(List[ElementId](groups))