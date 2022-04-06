# -*- coding: utf-8 -*-

import clr
clr.AddReference("PresentationCore")

from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.Exceptions import *

from pyrevit import forms
from dosymep_libs.bim4everyone import *

document = __revit__.ActiveUIDocument.Document
uiDocument = __revit__.ActiveUIDocument


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with forms.WarningBar(title="Выберите элемент связанного файла"):
        try:
            selectedElement = uiDocument.Selection.PickObject(ObjectType.LinkedElement, "Выберите элемент связанного файла.")
            linkedDocuments = [ documentInstance.GetLinkDocument() for documentInstance in FilteredElementCollector(document).OfClass(RevitLinkInstance) if documentInstance.Id == selectedElement.ElementId ]

            linkedDocument = next(iter(linkedDocuments), None)
            if linkedDocument:
                linkedElement = linkedDocument.GetElement(selectedElement.LinkedElementId)

                Clipboard.SetDataObject(linkedElement.Id.ToString())
                print "{title}.rvt ID: {elementId}".format(title=linkedDocument.Title, elementId=linkedElement.Id)
            else:
                forms.alert("Не был найден элемент в связанных файлах.", title="Сообщение")
        except OperationCanceledException:
            show_canceled_script_notification()


script_execute()