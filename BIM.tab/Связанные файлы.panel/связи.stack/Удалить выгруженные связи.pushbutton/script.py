# -*- coding: utf-8 -*-
import clr
clr.AddReference("System")

from System.Collections.Generic import *
from Autodesk.Revit.DB import *
from pyrevit import forms
from pyrevit import script
from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


document = __revit__.ActiveUIDocument.Document


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):

    linked_files = FilteredElementCollector(document)\
        .OfCategory(BuiltInCategory.OST_RvtLinks)\
        .WhereElementIsElementType()\
        .ToElementIds()

    linked_files = [ item for item in linked_files if RevitLinkType.IsLoaded(document, item) == False ]
    linked_file_names = [ document.GetElement(item) for item in linked_files]
    linked_file_names = [ item.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString() for item in linked_file_names ]

    if linked_file_names:
        linked_file_names.sort()
        result = forms.alert("Будут удалены следующие связанные файлы:\r\n - " + "\r\n - ".join(linked_file_names), title="Предупреждение", ok=False, yes=True, no=True)
        if not result:
            script.exit()

        with revit.Transaction("BIM: Удаление связанных файлов"):
            document.Delete(List[ElementId](linked_files))


script_execute()