# -*- coding: utf-8 -*-
import os.path as op
import os
import sys

from Autodesk.Revit.DB import  GroupType, FilteredElementCollector, Transaction, TransactionGroup
__doc__ = 'Выделяет все элементы, находящиеся в группах'
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

groupTypes = [x for x in FilteredElementCollector(doc).OfClass(GroupType).ToElements()]
tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()

for el in groupTypes:
	#print el.LookupParameter('Имя типа').AsString()
	#print el.Groups.Size
	if el.Groups.Size < 1:
		try:
			doc.Delete( el.Id )
		except:
			pass
		
#print dir(groupTypes[0])
#print len(groupTypes)

t.Commit()
tg.Assimilate()
'''
tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()

t.Commit()
tg.Assimilate()

MessageBox.Show('Готово!')
'''