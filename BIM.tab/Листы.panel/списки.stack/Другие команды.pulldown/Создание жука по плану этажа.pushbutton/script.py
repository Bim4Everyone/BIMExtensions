# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
import clr
clr.AddReference('ImageConverter')
clr.AddReference('MathNet.Numerics')
clr.AddReference('Xceed.Wpf.Toolkit')
clr.AddReference('System')
clr.AddReference("System.Windows.Forms")

from Autodesk.Revit.DB import *
from ImageConverter import UserControl1

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
view = __revit__.ActiveUIDocument.ActiveGraphicalView 
view = doc.ActiveView


legends = [x for x in FilteredElementCollector(doc).OfClass(View) if x.ViewType == ViewType.Legend]
legendsNames = [x.Name for x in legends]
legends = [x for x in legends if x.CanViewBeDuplicated(ViewDuplicateOption.Duplicate)]
baseLegend = legends[0]
legendName = "Попробовал"

walls = [x for x in FilteredElementCollector(doc, view.Id).OfClass(Wall).WhereElementIsNotElementType().ToElements()]
try:
	form = UserControl1()
except Exception as e:
	print e
if form.ShowDialog():

	scale = 1##.0/form.LegendScale
	legendScale = int(form.LegendScale)
	legendName = form.LegendName
	if legendName == "":
		legendName = "Легенда по виду " + view.Name
	while (legendName in legendsNames):
		legendName = legendName + " копия"

	transform = Transform.CreateTranslation(XYZ.Zero)
	transform = transform.ScaleBasis(scale) 

	tg = TransactionGroup(doc, "Update")
	tg.Start()
	t = Transaction(doc, "Update Sheet Parmeters")
	t.Start()		

	legendId = baseLegend.Duplicate(ViewDuplicateOption.Duplicate)
	legend = doc.GetElement(legendId)
	legend.Name = legendName

	legend.LookupParameter("Масштаб вида").Set(legendScale)

	for wall in walls:
		curve = wall.Location.Curve
		scaledCurve = curve.CreateTransformed(transform)

		
		if scaledCurve:
			doc.Create.NewDetailCurve(legend, scaledCurve)
			



	t.Commit()
	tg.Assimilate()

