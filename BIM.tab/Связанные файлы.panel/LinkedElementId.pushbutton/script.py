# -*- coding: utf-8 -*-
import clr
clr.AddReference("PresentationCore")

from System.Windows import *

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.Exceptions import *

from pyrevit import forms

document = __revit__.ActiveUIDocument.Document
uiDocument = __revit__.ActiveUIDocument

try:
    selectedElement = uiDocument.Selection.PickObject(ObjectType.LinkedElement, "Выберите элемент связанного файла.")
    linkedDocuments = [ documentInstance.GetLinkDocument() for documentInstance in FilteredElementCollector(document).OfClass(RevitLinkInstance) if documentInstance.Id == selectedElement.ElementId ]
    
    linkedDocument = next(iter(linkedDocuments), None)
    if linkedDocument:
        linkedElement = linkedDocument.GetElement(selectedElement.LinkedElementId)

        Clipboard.SetText(linkedElement.Id.ToString());
        print "{title}.rvt ID: {elementId}".format(title=linkedDocument.Title, elementId=linkedElement.Id)
    else:
        forms.alert("Не был найден элемент в связанных файлах.", title="Сообщение")
except OperationCanceledException:
    pass