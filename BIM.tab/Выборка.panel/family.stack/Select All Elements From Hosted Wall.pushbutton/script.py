# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import Wall, FilteredElementCollector, BuiltInCategory, BuiltInParameter, SpatialElementBoundaryOptions, Transaction, TransactionGroup, FamilyInstance, WallSweep

from pyrevit import revit, DB

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
error_collector = {}

__title__ = 'Элементы основы'
__context__ = 'Selection'
__doc__ = 'Выбреет все элементы, размещенные на выбранной основе'
selection =  __revit__.ActiveUIDocument.Selection.GetElementIds()
temp = FilteredElementCollector(doc).OfClass(FamilyInstance).ToElements()
wallsweeps = FilteredElementCollector(doc).OfClass(WallSweep).ToElements()

associate = []
walls = []
for elid in selection:
	el = doc.GetElement(elid) 
	if isinstance(el, Wall):
		walls.append(elid)
		associate.append(elid)

for t in temp:
	if t.Host:
		for wall in walls:
			if str(t.Host.Id) == wall.ToString():
				associate.append(t.Id)
				break
			
for ws in wallsweeps:
	for wall in walls:
		if wall in ws.GetHostIds():
			associate.append(ws.Id)
			break


selection = revit.get_selection()
selection.set_to(associate)
