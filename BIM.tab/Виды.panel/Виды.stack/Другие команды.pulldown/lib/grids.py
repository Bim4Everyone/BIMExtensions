# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *
from pyrevit import revit, EXEC_PARAMS


def switch_grids(transaction_name, datum_extent_type):
    doc = __revit__.ActiveUIDocument.Document
    view = __revit__.ActiveUIDocument.ActiveView

    grids = FilteredElementCollector(doc, view.Id) \
        .OfCategory(BuiltInCategory.OST_Grids) \
        .WhereElementIsNotElementType() \
        .ToElements()

    with revit.Transaction("BIM: " + transaction_name):
        for grid in grids:
            grid.SetDatumExtentType(DatumEnds.End1, view, datum_extent_type)
            grid.SetDatumExtentType(DatumEnds.End0, view, datum_extent_type)