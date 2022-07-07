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


def get_tags(doc, view):
    tags = FilteredElementCollector(doc, view.Id) \
        .OfClass(IndependentTag) \
        .WhereElementIsNotElementType()\
        .ToElements() # type: list[IndependentTag]

    for tag in tags:
        yield tag

    tags = FilteredElementCollector(doc, view.Id) \
        .OfClass(SpatialElementTag) \
        .WhereElementIsNotElementType() \
        .ToElements()  # type: list[SpatialElementTag]

    for tag in tags:
        yield tag


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    doc = __revit__.ActiveUIDocument.Document # type: Document
    view = __revit__.ActiveUIDocument.ActiveView # type: View

    tags = get_tags(doc, view)
    tags = [tag for tag in tags
            if not tag.TagText or tag.TagText == "0" or tag.TagText == "?"]

    with revit.Transaction("BIM: Удаление пустых марок"):
        for tag in tags:
            doc.Delete(tag.Id)


script_execute()
