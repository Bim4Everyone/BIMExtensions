# -*- coding: utf-8 -*-
import clr
clr.AddReference("System.Windows.Forms")

from System.IO import Path
from System.Windows.Forms import *

from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import *

from pyrevit import forms

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


selectedFiles = forms.pick_file(files_filter="Revit files (*.rvt)|*.rvt", multi_file=True, title="Выберите Revit файлы")
if selectedFiles:
    selectedFileNames = [ Path.GetFileName(value).lower() for value in list(selectedFiles) ]
    selectedFiles = dict(zip(selectedFileNames, selectedFiles))

    # Перезагружаем существующие связи
    ReloadExistingLinks(selectedFiles)
    
    # Добавляем оставшиеся связи
    CreateRevitLinks(selectedFiles)