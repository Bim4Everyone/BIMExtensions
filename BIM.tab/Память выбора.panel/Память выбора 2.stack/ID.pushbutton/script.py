# -*- coding: utf-8 -*-

import clr
clr.AddReference("PresentationCore")

from System.Windows import Clipboard

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

from pyrevit import forms
from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

document = __revit__.ActiveUIDocument.Document
uiDocument = __revit__.ActiveUIDocument


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    with forms.WarningBar(title="Выберите элемент связанного файла"):
        selected_element = uiDocument.Selection.PickObject(ObjectType.LinkedElement,
                                                           "Выберите элемент связанного файла")

        link_doc = document.GetElement(selected_element.ElementId)
        if not link_doc:
            forms.alert("Не был найден связанный документ.", exitscript=True)

        link_doc = link_doc.GetLinkDocument()
        link_element = link_doc.GetElement(selected_element.LinkedElementId)
        if not link_doc:
            forms.alert("Не был найден элемент в связанных файлах.", exitscript=True)

        Clipboard.SetText(str(link_element.Id))
        print "{title}.rvt ID: {elementId}".format(title=link_doc.Title, elementId=link_element.Id)


script_execute()
