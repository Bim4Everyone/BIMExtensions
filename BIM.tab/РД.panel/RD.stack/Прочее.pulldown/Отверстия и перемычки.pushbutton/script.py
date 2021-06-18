# -*- coding: utf-8 -*-
import os.path as op
import os
import sys

from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, Transaction, TransactionGroup,FamilySymbol
from pyrevit import revit, DB

#__title__ = 'Перенос данных(из Перемычки в Отверстие)'

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

#selection = [ doc.GetElement( elId ) for elId in __revit__.ActiveUIDocument.Selection.GetElementIds() ]

famSym = [x.Family.GetFamilySymbolIds() for x in FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements() if x.Family.Name == 'АР-Отверстие под инженерные системы'] 
#for sym in famSym:
#	print sym
emptyList = []
tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()
for fam in famSym:
	for symid in fam:
		selection = DB.FilteredElementCollector(revit.doc)\
					   .WherePasses(DB.FamilyInstanceFilter(revit.doc, symid))\
					   .ToElements()
		for el in selection:
			temp = ['','','']
			ex_par = el.LookupParameter('Наличие Перемычки').AsInteger()
			if ex_par:
				subelId = el.GetSubComponentIds()
				subel = doc.GetElement(subelId[0])

				temp[0] = int(subel.LookupParameter('Speech_Размер_Ширина').AsDouble()*304.8)

				subel_type = doc.GetElement(subel.GetTypeId())

				temp[1] = int(subel_type.LookupParameter('Speech_Размер_Глубина').AsDouble()*304.8)
				temp[2] = subel_type.LookupParameter('Маркировка типоразмера').AsString()
					
				pSet = el.Parameters
				el.LookupParameter('Перемычка Длина').Set(str(temp[0]))
				el.LookupParameter('Перемычка Глубина').Set(str(temp[1]))
				if temp[2] is None:
					emptyList.append(el.Id)
				else:
					el.LookupParameter('Перемычка Класс').Set(temp[2])

			else:
				pSet = el.Parameters
				el.LookupParameter('Перемычка Длина').Set(str(temp[0]))
				el.LookupParameter('Перемычка Глубина').Set(str(temp[1]))
				el.LookupParameter('Перемычка Класс').Set(temp[2])
			#print temp

t.Commit()
tg.Assimilate()
'''
el = selection[0]
pSet = el.Parameters
for param in pSet:
	name = param.Definition.Name
	print name
	hex_name = [elem.encode("hex") for elem in name]
	print [elem.encode("hex") for elem in '������� ����']
	print hex_name
	bin_name = ' '.join(format(ord(x), 'b') for x in name)
	print ' '.join(format(ord(x), 'b') for x in '������� ����')
	print bin_name
	break
'''
if emptyList:
	print "Не заполнен параметр 'Маркировка типоразмера'"
	for index in emptyList:
		print index