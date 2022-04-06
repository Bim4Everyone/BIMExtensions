# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep.Revit
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from System.Diagnostics import Process

from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    Process.Start(r"T:\Проектный институт\Отдел стандартизации BIM и RD\BIM-Ресурсы\3-00_Семейства Общие\Подписи")


script_execute()