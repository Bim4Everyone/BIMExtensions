# -*- coding: utf-8 -*-
import clr
clr.AddReference('System')
clr.AddReference('ImageConverter')
clr.AddReference('MathNet.Numerics')
clr.AddReference('Xceed.Wpf.Toolkit')
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
from System.Collections.Generic import List 
from Autodesk.Revit.DB import Options, FilledRegionType, FilteredElementCollector, ImportInstance, BuiltInCategory, FamilySymbol, TransactionGroup, Transaction
from ImageConverter import ImageController

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
view = __revit__.ActiveUIDocument.ActiveGraphicalView 
view = doc.ActiveView

selection = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]

LAYER_NAME = 'клинкер'
if len(selection) < 1:
    MessageBox.Show("Выберите импортированный объект DWG!")
    raise SystemExit(1)
    
dwg_link_instance = selection[0]


if isinstance(dwg_link_instance, ImportInstance):
    active_view = doc.ActiveView
    geo_opt = Options()

    geo_opt.ComputeReferences = True
    geo_opt.IncludeNonVisibleObjects = True
    geo_opt.View = active_view
    defaultFilledRegionType = None
    geometry = dwg_link_instance.get_Geometry(geo_opt)
    


    imageController = ImageController(geometry, view, True)

    if imageController.ImagePicked:
        tg = TransactionGroup(doc, "Update")
        tg.Start()
        t = Transaction(doc, "Update Sheet Parmeters")
        t.Start()
        
        imageController.CreateGeometry()

        t.Commit()
        tg.Assimilate()
else:
    MessageBox.Show("Выберите импортированный объект DWG!")