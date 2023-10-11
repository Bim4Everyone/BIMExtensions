# -*- coding: utf-8 -*-

import clr

clr.AddReference("PresentationCore")

from System.Windows import Clipboard

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *
from Autodesk.Revit.Exceptions import *

from pyrevit import forms
from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

document = __revit__.ActiveUIDocument.Document
uiDocument = __revit__.ActiveUIDocument


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with forms.WarningBar(title="Выберите элемент связанного файла"):
        try:
            selected_element = uiDocument.Selection.PickObject(ObjectType.LinkedElement,
                                                               "Выберите элемент связанного файла.")

            linked_documents = [documentInstance.GetLinkDocument() for documentInstance in
                                FilteredElementCollector(document).OfClass(RevitLinkInstance)
                                if documentInstance.Id == selected_element.ElementId]

            linked_document = next(iter(linked_documents), None)
            if linked_document:
                linked_element = linked_document.GetElement(selected_element.LinkedElementId)

                Clipboard.SetDataObject(linked_element.Id.ToString())
                print "{title}.rvt ID: {elementId}".format(title=linked_document.Title, elementId=linked_element.Id)
            else:
                forms.alert("Не был найден элемент в связанных файлах.", title="Сообщение")
        except OperationCanceledException:
            show_canceled_script_notification()


script_execute()
