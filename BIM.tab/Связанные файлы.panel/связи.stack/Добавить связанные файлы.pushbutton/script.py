# -*- coding: utf-8 -*-
import clr
clr.AddReference("System.Windows.Forms")

from System.IO import Path
from System.Windows.Forms import *

from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import *

from pyrevit import forms
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

document = __revit__.ActiveUIDocument.Document


def ReloadExistingLinks(selectedFiles):
    linkedFiles = FilteredElementCollector(document)\
        .OfCategory(BuiltInCategory.OST_RvtLinks)\
        .WhereElementIsElementType()\
        .ToElements()

    for linkFile in linkedFiles:
        linkName = linkFile.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_NAME).AsString()
        linkFilePath = selectedFiles.pop(linkName.lower(), None)
       
        if linkFilePath:
            linkFileModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(linkFilePath)
            linkFile.LoadFrom(linkFileModelPath, WorksetConfiguration())
            

def CreateRevitLinks(selectedFiles):
    errorList = []
    with TransactionGroup(document) as transactionGroup:
        transactionGroup.Start("Связывание файлов")

        for fileName, fileNamePath in selectedFiles.items():
            with Transaction(document) as transaction:
                transaction.Start("Связывание файла " + Path.GetFileName(fileNamePath))
                
                try:                    
                    linkFile = ModelPathUtils.ConvertUserVisiblePathToModelPath(fileNamePath)
                    
                    linkOptions = RevitLinkOptions(True)
                    linkLoadResult = RevitLinkType.Create(document, linkFile, linkOptions)

                    revitLinkInstance = RevitLinkInstance.Create(document, linkLoadResult.ElementId, ImportPlacement.Shared)
                    
                    revitLinkType = document.GetElement(revitLinkInstance.GetTypeId())
                    revitLinkType.Parameter[BuiltInParameter.WALL_ATTR_ROOM_BOUNDING].Set(1)

                    transaction.Commit()
                except InvalidOperationException:
                    transaction.RollBack()
                    errorList.append(Path.GetFileName(fileNamePath))
        
        transactionGroup.Assimilate()

    if errorList:
        forms.alert("Файлы с другой системой координат:\r\n - " + "\r\n - ".join(errorList), title="Предупреждение")


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selected_files = forms.pick_file(files_filter="Revit files (*.rvt)|*.rvt", multi_file=True, title="Выберите Revit файлы")
    if selected_files:
        selected_file_names = [ Path.GetFileName(value).lower() for value in list(selected_files) ]
        selected_files = dict(zip(selected_file_names, selected_files))

        # Перезагружаем существующие связи
        ReloadExistingLinks(selected_files)

        # Добавляем оставшиеся связи
        CreateRevitLinks(selected_files)
        show_executed_script_notification()


script_execute()