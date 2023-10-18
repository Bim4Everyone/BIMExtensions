# -*- coding: utf-8 -*-
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import revit
from pyrevit import framework
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

document = __revit__.ActiveUIDocument.Document


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    linked_files = FilteredElementCollector(document) \
        .OfCategory(BuiltInCategory.OST_RvtLinks) \
        .WhereElementIsElementType() \
        .ToElementIds()

    linked_files = [item for item in linked_files if not RevitLinkType.IsLoaded(document, item)]
    linked_file_names = [document.GetElement(item) for item in linked_files]
    linked_file_names = [item.GetParamValue(BuiltInParameter.ALL_MODEL_TYPE_NAME) for item in linked_file_names]

    if linked_file_names:
        linked_file_names.sort()
        result = forms.alert("Будут удалены следующие связанные файлы:\r\n - " + "\r\n - ".join(linked_file_names),
                             ok=False, yes=True, no=True)
        if not result:
            script.exit()

        with revit.Transaction("BIM: Удаление связанных файлов"):
            revit.delete.delete_elements(linked_files, document)


script_execute()
