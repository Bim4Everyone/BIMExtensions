# -*- coding: utf-8 -*-
import clr
clr.AddReference("OpenMcdf.dll")
clr.AddReference("dosymep.Revit.dll")

from dosymep.Revit import *
from pyrevit import forms
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selected_files = forms.pick_file(files_filter="Revit (*.rvt)|*.rvt", multi_file=True, title="Выберите Revit файлы")
    if not selected_files:
        script.exit()

    DocumentExtensions.UnloadAllLinks(selected_files)


script_execute()