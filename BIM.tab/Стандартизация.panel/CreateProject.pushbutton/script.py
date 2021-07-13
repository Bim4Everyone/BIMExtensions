# -*- coding: utf-8 -*-

import clr
clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import *
from Autodesk.Revit.DB import *

templateFilePath = "T:\\Проектный институт\\Отдел стандартизации BIM и RD\\BIM-Ресурсы\\5-Надстройки\\Шаблоны и настройки\\Test\\Пустой.rte"

application = __revit__.Application

def GetFilePath():
    with SaveFileDialog() as dialog:
        dialog.RestoreDirectory = True
        dialog.Filter = "Revit files (*.rvt)|*.rvt"
        
        if dialog.ShowDialog() == DialogResult.OK:
            return dialog.FileName

    return None


fileName = GetFilePath()
if(fileName):
    document = application.NewProjectDocument(templateFilePath)
    document.SaveAs(fileName, SaveAsOptions(OverwriteExistingFile=True))
    document.Close(False)

    __revit__.OpenAndActivateDocument(fileName)