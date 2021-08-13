# -*- coding: utf-8 -*-
from pyrevit.framework import List
from pyrevit import revit, DB


cl = DB.FilteredElementCollector(revit.doc, revit.activeview.Id)\
       .WhereElementIsNotElementType()\
       .ToElementIds()

matchlist = []
selCatList = set()

selection = revit.get_selection()

for el in selection:
    try:
        selCatList.add(el.Category.Name)
    except Exception:
        continue

for elid in cl:
    el = revit.doc.GetElement(elid)
    try:
        # if el.ViewSpecific and ( el.Category.Name in selCatList):
        if el.Category.Name in selCatList:
            matchlist.append(elid)
    except Exception:
        continue

selSet = []
for elid in matchlist:
    selSet.append(elid)

selection.set_to(selSet)
revit.uidoc.RefreshActiveView()
