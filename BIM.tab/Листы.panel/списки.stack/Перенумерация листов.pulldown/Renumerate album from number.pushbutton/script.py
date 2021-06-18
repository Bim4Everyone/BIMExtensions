# -*- coding: utf-8 -*-
#импорты стандартных библиотек 
from operator import itemgetter
import clr, sys
clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import MessageBox
from System.Drawing import Point
#импорт из библиотеки Autodesk
from Autodesk.Revit.DB import ViewSheet, FilteredElementCollector, BuiltInCategory, BuiltInParameter, Transaction, TransactionGroup
#ипорт из библиотеки pyRevit
from pyrevit import forms
from pyrevit.framework import Controls
from pyrevit.forms import SelectFromCheckBoxes, SelectFromList
from pyrevit.coreutils import Timer
timer=Timer()

import os

pylocalpath = os.path.join(os.getenv('APPDATA'), 'pyRevit\extensions\lib')
sys.path.append(pylocalpath) 

#Speech библиотека
from pySpeech.Forms import InputFormNumber
from pySpeech import alert, AlbumFilter, ClearString, GetNumber, FormNum

__title__ = 'Перенумерация\nс номера'
__doc__ = 'Перенумеровать лист/группу листов начиная с введенного номера. '
		
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selection = [ doc.GetElement( elId ) for elId in __revit__.ActiveUIDocument.Selection.GetElementIds() ]
#количество выбранных листов 
count = len(selection) 


import sys


if not count:
	alert("Не выбран лист")
elif all(isinstance(n, ViewSheet ) for n in selection):
	return_options =InputFormNumber.show([], title='Введите номер', button_name='Ок', width=250, height=130)
	#шаг перемещения
	inpos = int(return_options)
	#список из листов учавствующих в перестановке
	SheetsCurrentAlbumList = []
	#список, на основе кторого будет сортировка
	SortingList = []
	#список, на основе которого будет происходить перенумерование листов
	SheetsComposeAlbumList = []
	#пара значений, будет хранить в себе номер крайнего листа среди рабочих
	if selection is None: 
		alert("Выборка пуста")
		sys.exit()
	parameter = selection[0].LookupParameter("Номер листа")
	if parameter is None: 
		alert("Номер листа отсутствует")
		sys.exit()
	prop = [GetNumber(parameter.AsString().split(".")), ""]
	if prop is None: 
		alert("Номер листа не найден")
		sys.exit()
	parameter = selection[0].LookupParameter("Альбом")
	if parameter is None: 
		alert("Альбом отсутствует")
		sys.exit()#название альбома
	album = parameter.AsString()
	if album is None: 
		alert("Альбом не найден")
		sys.exit()
	#все листы проекта
	SheetsBaseList = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).ToElements()
	
	for sheet in SheetsBaseList:
		if  AlbumFilter(sheet, album):
			good = True
			number = sheet.LookupParameter("Номер листа").AsString()
			for sel in selection:
				if sel.LookupParameter("Номер листа").AsString()==number:
					good = False
					break
			if good:
				SheetsCurrentAlbumList.append(number)
	
	#составляем список из троек значений, [пригодный для сортировки номер, сам лист, номер листа в строковой форме]
	for sheet in selection:
		NumberString = sheet.LookupParameter("Номер листа").AsString()
		NumberList = GetNumber(NumberString.split("."))
		SortingList.append([NumberList, sheet])
	

	#сортируем список
	selection=[x[1] for x in sorted(SortingList, key=itemgetter(0))]
	
	NewNumList = FormNum(album, inpos, count)
	
	good = True
	for num in SheetsCurrentAlbumList:
		if good:
			for sheet in NewNumList:
				if sheet==num:
					good = False
					break
		else:
			break
	
	if not good:
		alert("Номер занят!")
	else:
		tg = TransactionGroup(doc, "Update")
		tg.Start()
		t = Transaction(doc, "Update Sheet Parmeters")
		t.Start()
		
		
		for idx, sheet in enumerate(selection):
			sheet.LookupParameter("Номер листа").Set(str(idx)+"+Temp")
		for idx, sheet in enumerate(selection):
			sheet.LookupParameter("Номер листа").Set(NewNumList[idx])
			sheet.LookupParameter("Номер листа фактический").Set(str(inpos+idx))

		t.Commit()

		tg.Assimilate()
	
	'''
#--------------------------------------------------------------#
	print "Количество листов: "+str(len(SheetsBaseList))	
	print "Выбрано: "+str(count)
	print "Опорный номер: "+prop[1]
	print "Альбом: "+str(album)

	for sheet in SheetsCurrentAlbumList:
		print sheet
	print "Размер альбома: "+str(len(SheetsCurrentAlbumList))
	print "Время: "+str(timer.get_time())


	
#--------------------------------------------------------------#
	
	'''
	
	
