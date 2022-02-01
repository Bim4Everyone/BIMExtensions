# -*- coding: utf-8 -*-
import clr
clr.AddReference("OpenMcdf.dll")
clr.AddReference("dosymep.Revit.dll")

from dosymep.Revit import *
from pyrevit import forms

selectedFiles = forms.pick_file(files_filter="Revit files (*.rvt)|*.rvt", multi_file=True, title="Выберите Revit файлы")
if selectedFiles:
    DocumentExtensions.UnloadAllLinks(selectedFiles)