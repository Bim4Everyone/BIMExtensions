# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit.DB import *
from pyrevit import revit, EXEC_PARAMS

from dosymep_libs.bim4everyone import log_plugin
from dosymep_libs.simple_services import notification

from grids import switch_grids


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
	switch_grids("Отключение 3D оси", DatumExtentType.ViewSpecific)


script_execute()