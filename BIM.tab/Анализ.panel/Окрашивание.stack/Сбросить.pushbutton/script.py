# -*- coding: utf-8 -*-

from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    application = __revit__.Application
    document = __revit__.ActiveUIDocument.Document
    activeView = document.ActiveView

    elementsInView = FilteredElementCollector(document, activeView.Id)\
        .WhereElementIsNotElementType()\
        .ToElements()

    settings = OverrideGraphicSettings()

    with Transaction(document) as transaction:
        transaction.Start("Сброс переопределения графики")
        for elem in elementsInView:
            activeView.SetElementOverrides(elem.Id, settings)
        transaction.Commit()


script_execute()