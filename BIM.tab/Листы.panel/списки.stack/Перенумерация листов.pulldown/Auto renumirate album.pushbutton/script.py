# -*- coding: utf-8 -*-
#импорты стандартных библиотек 
from operator import itemgetter
import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
#импорт из библиотеки Autodesk
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, Transaction, TransactionGroup
#ипорт из библиотеки pyRevit
from pyrevit.coreutils import Timer
timer=Timer()

__title__ = 'Автонумерация'

__doc__ = 'Полностью перенумеровывает альбом, лист которого выделен.\n'\
		'Нумерация начинается с 1.'



#Всплывающее сообщение
def alert(msg):
	MessageBox.Show(msg)
#Проверка на соответствие выбранному альбому
def AlbumFilter(x, album):
	if x is None: return 0
	if album is None: return 0
	parameter = x.LookupParameter("Альбом")
	if parameter is None: return 0
	
	if parameter.AsString()==album:
		return 1
	else:
		return 0
#Получение числа из строки
def ClearString(msg): 
	number=""
	for i in range(len(msg)):
		if msg[i].isdigit():
			number+=msg[i]
	return number
#Получаем число для сравнения из строки вида "str-num.num"
def GetNumber(list):
	first = ClearString(list[0])
	first = float(first)
	if len(list)>1:
		second = int(list[1])
		if second>0:
			first+=1-0.5/second
		else:
			first+=0.6
	return first
	
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selection = [ doc.GetElement( elId ) for elId in __revit__.ActiveUIDocument.Selection.GetElementIds() ]
count = len(selection)

import sys

if not count:
	alert("Не выбран лист")
else:
	#список, на основе кторого будет сортировка
	SortingList = []
	if selection is None: 
		alert("Выборка пуста")
		sys.exit()
	#название альбома
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
	
	#составляем список из троек значений, [пригодный для сортировки номер, сам лист, номер листа в строковой форме]
	for sheet in SheetsBaseList:
		if  AlbumFilter(sheet, album):
			NumberString = sheet.LookupParameter("Номер листа").AsString()
			NumberList = GetNumber(NumberString.split("."))
			SortingList.append([NumberList, sheet])
			
	#сортируем список
	SheetsBaseList= [ sheet[1] for sheet in sorted(SortingList, key=itemgetter(0)) ]

	
	
	tg = TransactionGroup(doc, "Update")
	tg.Start()
	t = Transaction(doc, "Update Sheet Parmeters")
	t.Start()
	
	
	for idx, sheet in enumerate(SheetsBaseList):
		sheet.LookupParameter("Номер листа").Set(str(idx)+"+Temp")
	for idx, sheet in enumerate(SheetsBaseList):
		sheet.LookupParameter("Номер листа").Set(album+"-"+str(idx+1))
		sheet.LookupParameter("Номер листа фактический").Set(str(idx+1))

	t.Commit()

	tg.Assimilate()

	'''
#--------------------------------------------------------------#
	print "Количество листов: "+str(len(SheetsBaseList))	
	print "Выбрано: "+str(count)
	print "Опорный номер: "+prop[1]
	print "Альбом: "+str(album)
	for sheet in SheetsCurrentAlbumList:
		print sheet[1].LookupParameter("Номер листа").AsString()+' '+sheet[1].LookupParameter("Имя листа").AsString()
	print "Размер альбома: "+str(len(SheetsCurrentAlbumList))
	print "Время: "+str(timer.get_time())
	for sheet in SheetsComposeAlbumList:
		print sheet[1]
	
	
#--------------------------------------------------------------#
	
	'''
	
	
