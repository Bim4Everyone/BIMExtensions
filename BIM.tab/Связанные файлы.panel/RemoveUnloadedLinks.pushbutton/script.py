# -*- coding: utf-8 -*-
import clr
clr.AddReference("System")

from System.Collections.Generic import *
from Autodesk.Revit.DB import *
from pyrevit import forms

document = __revit__.ActiveUIDocument.Document

linkedFiles = FilteredElementCollector(document)\
    .OfCategory(BuiltInCategory.OST_RvtLinks)\
    .WhereElementIsElementType()\
    .ToElementIds()

linkedFiles = [ item for item in linkedFiles if RevitLinkType.IsLoaded(document, item) == False ]
linkedFileNames = [ document.GetElement(item) for item in linkedFiles]
linkedFileNames = [ item.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() for item in linkedFileNames ]

if linkedFileNames:
    linkedFileNames.sort()
    result = forms.alert("Будут удалены следующие связанные файлы:\r\n - " + "\r\n - ".join(linkedFileNames), title="Предупреждение", ok=False, yes=True, no=True)

    if result:
        with Transaction(document, "Удаление связанных файлов") as transaction:
            transaction.Start()

            document.Delete(List[ElementId](linkedFiles))
        
            transaction.Commit()