# -*- coding: utf-8 -*-

from System import InvalidOperationException
from Autodesk.Revit.DB import FilteredElementCollector, Transaction, BuiltInParameter, ViewType

from pyrevit import forms

application = __revit__.Application
document = __revit__.ActiveUIDocument.Document

schedule = document.ActiveView
if schedule.ViewType != ViewType.Schedule:
    raise InvalidOperationException("Данная операция доступна только для спецификаций.")

elements = FilteredElementCollector(schedule.Document, schedule.Id).ToElements()

tran = Transaction(schedule.Document, "Обновление спецификации")
tran.Start()

for element in elements:  
    typeParameter = element.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM)
    nameParameter = element.get_Parameter(BuiltInParameter.RVT_LINK_INSTANCE_NAME)
    
    if typeParameter is not None and nameParameter is not None:
        nameParameter.Set(typeParameter.AsValueString())

tran.Commit()