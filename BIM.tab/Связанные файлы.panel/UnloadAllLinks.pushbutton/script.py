# -*- coding: utf-8 -*-
import clr
clr.AddReference("dosymep.Revit.dll")

from pyrevit import forms

selectedFiles = forms.pick_file(files_filter="Revit files (*.rvt)|*.rvt", multi_file=True, title="Выберите Revit файлы")
if selectedFiles:
    DocumentExtensions.UnloadRevitLinks(selectedFiles)