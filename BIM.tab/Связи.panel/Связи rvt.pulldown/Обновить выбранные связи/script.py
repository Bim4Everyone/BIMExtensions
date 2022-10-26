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
    pass


script_execute()