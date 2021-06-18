# -*- coding: utf-8 -*-
class Environment:

	#std
	import sys
	import traceback    
	import unicodedata
	import os
	#clr    
	import clr
	clr.AddReference('System')
	clr.AddReference('System.IO')    
	clr.AddReference('PresentationCore')
	clr.AddReference("PresentationFramework")
	clr.AddReference("System.Windows")
	clr.AddReference("System.Xaml")
	clr.AddReference("WindowsBase")


	from System.Windows import MessageBox
	#
	@classmethod
	def Exit(cls): cls.sys.exit("Exit current environment!")

	@classmethod
	def GetLastError(cls):
		info = cls.sys.exc_info()
		infos = ''.join(cls.traceback.format_exception(info[0], info[1], info[2]))
		return infos

	@classmethod
	def GetString(cls, array, delim):
		if not isinstance(array, list): return ""
		if array is None: return ""
		if delim is None: delim = ""
		strn = ""
		i = 1
		l = len(array)
		strn += array[0]
		while i < l:
			elemt = array[i]
			if elemt is None: elemt = ""
			if type(elemt) is not str: elemt = str(elemt)
			if strn is not "": strn += delim
			strn += elemt
			i += 1
		return strn

	@classmethod
	def Message(cls, msg):
		if not msg is None and type(msg) == str:
			cls.MessageBox.Show(msg)
		else:
			cls.MessageBox.Show("Empty argument or non string error message!")

	@classmethod
	def SafeCall(cls, code, tail, show):        
		back = None
		flag = 0
		info = "empty"
		if not code is None:
			try:
				back = code(tail)
				flag = 1
			except:
				exc_t, exc_v, exc_i = sys.exc_info()           
				info = ''.join(cls.traceback.format_exception(exc_t, exc_v, exc_i))         
				if show: cls.Message(info)         
		return [flag, back, info]


#импорты стандартных библиотек 
from operator import itemgetter
import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
#импорт из библиотеки Autodesk
from Autodesk.Revit.DB import ViewSheet, FilteredElementCollector, BuiltInCategory, BuiltInParameter, Transaction, TransactionGroup
#ипорт из библиотеки pyRevit
from pyrevit.coreutils import Timer
timer=Timer()


import sys

class localnames:
	sheet_set1 = "ФОП_Комплект Чертежей"
	sheet_set2 = "ADSK_Комплект чертежей"
	sheet_set3 = "Орг.КомплектЧертежей"
	sheet_num1 = "ФОП_Номер Листа Фактический"
	sheet_num2 = "Ш.НомерЛиста"



__doc__ = 'Перемещает лист/группу листов вниз в списке альбома.'
#Всплывающее сообщение
def alert(msg):
	MessageBox.Show(msg)
#Проверка на соответствие выбранному альбому
def AlbumFilter(x, album):
	if x is None: return 0
	if album is None: return 0
	parameter = x.LookupParameter(localnames.sheet_set1)
	if parameter is None: parameter = x.LookupParameter(localnames.sheet_set2)
	if parameter is None: parameter = x.LookupParameter(localnames.sheet_set3)
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
		first+=1-0.5/float(list[1]) if list[1]>0 else 0.6
	return first
	
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selection = [ doc.GetElement( elId ) for elId in __revit__.ActiveUIDocument.Selection.GetElementIds() ]
#количество выбранных листов 
count = len(selection)
#шаг перемещения
step = -7
#направление перемещения
direction = 1 if step>0 else -1 
step = abs(step)

import sys

if not count:
	alert("Не выбран лист")
elif all(isinstance(n, ViewSheet ) for n in selection):
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
	parameter = selection[0].LookupParameter(localnames.sheet_num1)
	if parameter is None: parameter = selection[0].LookupParameter(localnames.sheet_num2)
	if parameter is None: 
		alert("Параметр 'Номер листа' отсутствует!")
		sys.exit()
	paramstr = parameter.AsString()
	if paramstr is None: 
		alert("Значение параметра 'Номер листа' не найдено!")
		sys.exit()
	prop = [GetNumber(paramstr.split(".")), ""]
	if prop is None: 
		alert("Значение параметра 'Номер листа' не найдено!")
		sys.exit()
	#пара значений, будет хранить в себе номер крайнего листа среди рабочих
	parameter = selection[0].LookupParameter(localnames.sheet_set1)
	if parameter is None: parameter = selection[0].LookupParameter(localnames.sheet_set2)
	if parameter is None: parameter = selection[0].LookupParameter(localnames.sheet_set3)
	if parameter is None: 
		alert("Параметр 'Комплект' отсутствует!")
		sys.exit()
	#название альбома
	album = parameter.AsString()
	if album is None: 
		alert("Значение параметра 'Комплект' не найдено!")
		sys.exit()	#все листы проекта

	#все листы проекта
	SheetsBaseList = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Sheets).ToElements()
	
	#составляем список из троек значений, [пригодный для сортировки номер, сам лист, номер листа в строковой форме]
	for sheet in SheetsBaseList:
		if  AlbumFilter(sheet, album):
			nparam = sheet.LookupParameter(localnames.sheet_num1)
			if nparam is None: nparam = sheet.LookupParameter(localnames.sheet_num2)
			if nparam is None: continue
			NumberString = nparam.AsString()
			if NumberString is None: continue
			NumberList = GetNumber(NumberString.split("."))
			SortingList.append([NumberList, sheet, NumberString])
			
	#сортируем список
	SheetsCurrentAlbumList=sorted(SortingList, key=itemgetter(0))
	#если перемещаем вниз, то нужно перевернуть список
	if direction<0:
		SheetsCurrentAlbumList=SheetsCurrentAlbumList[::-1]
	
	#находим опорное значение из выбранных листов, при уменьшении номера ищем большее, иначе - меньшее
	for sheet in selection:
		nparam = sheet.LookupParameter(localnames.sheet_num1)
		if nparam is None: nparam = sheet.LookupParameter(localnames.sheet_num2)
		if nparam is None: continue
		NumberString = nparam.AsString()
		if NumberString is None: continue
		NumberInt = GetNumber(NumberString.split("."))
		if direction*prop[0]<direction*NumberInt:
			prop[0] = NumberInt
			prop[1] = NumberString
	
	#идем с конца списка, удаляя все листы стоящие после наших выбранных
	for sheet in SheetsCurrentAlbumList[::-1]:
		if direction*prop[0]<direction*sheet[0]:
			SheetsCurrentAlbumList.pop()
		else:
			break
	
	#удаляем листы стоящие перед выбранными плюс step листов
	while len(SheetsCurrentAlbumList)>count+step:
		SheetsCurrentAlbumList.pop(0)
	
	#теперь размер списка постоянный, запишем его длину
	n=len(SheetsCurrentAlbumList)
	
	#кладем листы в нужном порядке первым значением из пары, второе значение это номера с неизменным порядком
	for i in range(n-count,n):
		element = [SheetsCurrentAlbumList[i][1],SheetsCurrentAlbumList[i-(n-count)][2]]
		SheetsComposeAlbumList.append(element)
	for i in range(n-count):
		element = [SheetsCurrentAlbumList[i][1], SheetsCurrentAlbumList[i+count][2]]
		SheetsComposeAlbumList.append(element)
	
	tg = TransactionGroup(doc, "Update")
	tg.Start()
	t = Transaction(doc, "Update Sheet Parmeters")
	t.Start()
	
	
	try:
		for idx, sheet in enumerate(SheetsComposeAlbumList):
			lparam = sheet[0].LookupParameter(localnames.sheet_num1)
			if lparam is None: lparam = sheet[0].LookupParameter(localnames.sheet_num2)
			if lparam is None: continue
			lparam.Set(str(idx)+"+Temp")
		for sheet in SheetsComposeAlbumList:
			lparam = sheet[0].LookupParameter(localnames.sheet_num1)
			if lparam is None: lparam = sheet[0].LookupParameter(localnames.sheet_num2)
			if lparam is None: continue
			lparam.Set(sheet[1])
	except:
		info = Environment.GetLastError()
		Environment.Message(info)


	t.Commit()

	tg.Assimilate()

	Environment.Message("Готово!")

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
	
	
