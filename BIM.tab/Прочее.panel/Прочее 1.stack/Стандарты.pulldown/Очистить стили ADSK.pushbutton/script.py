# -*- coding: utf-8 -*-
import os
import clr
import datetime
from System.Collections.Generic import *

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep

clr.AddReference("ClosedXML.dll")
import ClosedXML.Excel

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep_libs.bim4everyone import *

import pyevent
from pyrevit import EXEC_PARAMS, revit
from pyrevit.forms import *
from pyrevit import script

import Autodesk.Revit.DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application





@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Здравствуйте!")



    '''

    line = doc.GetElement(ElementId(11720162))

    style = line.LineStyle

    print(style)
    print(style.Name)

    cat = style.GraphicsStyleCategory

    print(cat.Name)

    elems =  cat.GetElements()








    elems = (FilteredElementCollector(doc)
             .OfClass(GraphicsStyle)
             .ToElements())

    print(str(len(elems)))

    for_del = []

    for elem in elems:
        if 'ADSK' in elem.Name:
            for_del.append(elem)

    with revit.Transaction("BIM: Заполнение параметров классификатора"):
        for elem in for_del:
            try:
                doc.Delete(elem.Id)
            except:
                print(elem.Name)
                
                '''









script_execute()
