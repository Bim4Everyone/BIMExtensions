# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.Exceptions import *

document = __revit__.ActiveUIDocument.Document
uiDocument = __revit__.ActiveUIDocument

try:
    selectedElement = uiDocument.Selection.PickObject(ObjectType.LinkedElement, "Выберите элемент связанного файла.")
    linkedDocuments = [ documentInstance.GetLinkDocument() for documentInstance in FilteredElementCollector(document).OfClass(RevitLinkInstance) if documentInstance.Id == selectedElement.ElementId ]
    
    linkedDocument = next(iter(linkedDocuments), None)
    if linkedDocument:
        linkedElement = linkedDocument.GetElement(selectedElement.LinkedElementId)
        print "{title} {elementId}".format(title=linkedDocument.Title, elementId=linkedElement.Id)
    else:
        print "Не был найден элемент в связанных файлах."
except OperationCanceledException:
    pass