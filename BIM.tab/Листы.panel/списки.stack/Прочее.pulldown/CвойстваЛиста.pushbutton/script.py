# -*- coding: utf-8 -*-
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


Environment.Message("Данная команда временно недоступна!")
Environment.Exit()

import os.path as op
import os
import re
import sys
import clr
import math
import collections
from shutil import copyfile
from pyrevit import forms
from pyrevit import revit

from pyrevit.framework import Controls
clr.AddReference('System')
clr.AddReference('System.IO')
clr.AddReference("System.Windows.Forms") 
#clr.AddReference("System.Windows.Controls.Primitives")
#from System.Windows.Controls.Primitives import Thumb
from System.IO import FileInfo
from System.Windows.Forms import MessageBox, SaveFileDialog, DialogResult
from System.Collections.Generic import List
from Autodesk.Revit.DB import DetailNurbSpline, CurveElement, ElementTransformUtils, \
							  DetailLine, View, ViewDuplicateOption, XYZ, LocationPoint, \
							  TransactionGroup, Transaction, FilteredElementCollector, \
							  ElementId, BuiltInCategory, FamilyInstance, ViewDuplicateOption, \
							  ViewSheet, FamilySymbol, Viewport, DetailEllipse, DetailArc, TextNote, \
							  ScheduleSheetInstance,Element
from Autodesk.Revit.Creation import ItemFactoryBase
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.UI import RevitCommandId, PostableCommand

import re
import os.path as op
import codecs
from collections import namedtuple

from pyrevit import HOST_APP
from pyrevit import USER_DESKTOP
from pyrevit import framework
from pyrevit.framework import Windows, Drawing, Forms
from pyrevit import coreutils
from pyrevit import forms
from pyrevit import revit, DB
from pyrevit import script


__title__ = 'Свойства Листа'
class SheetOption(object):
	def __init__(self, obj,sheet_name,list_name,list_familyName,
					familyInstance_Id,familySymbol_Id,speech_format,foreGround,state,visibility):
		self.state = state
		self.name = sheet_name
		self.familyName = "{} - {}".format(list_name,list_familyName)
		self.id = obj.Id
		self.familyInstance_Id = familyInstance_Id
		self.familySymbol_Id = familySymbol_Id
		self.speech_format = speech_format
		self.Number = obj.SheetNumber #obj.SheetNumber
		self.speech_album = ""
		self.speech_album_param = obj.LookupParameter('Альбом')
		if self.speech_album_param is None:
			Environment.Message("Параметр 'Альбом' отсутствует!")
		else:
			self.speech_album = self.speech_album_param.AsString()
		self.obj = obj
		self.foreGround = foreGround
		self.Visib = visibility
		def __nonzero__(self):
			return self.state
		def __str__(self):
			return self.name
class SymbolOption(object):
	def __init__(self, obj):
		self.familyName = '{} '.format(obj.FamilyName)
		self.Name = '{}	'.format(Element.Name.GetValue(obj))
		self.id = obj.Id
		self.obj = obj
def sort_fun(str):
	list = [int(x) for x in re.findall(r'\b\d+\b', str)]
	return list[0]
class SymbolStyle(object):
	def __init__(self, obj, symbolName,symbolColor, state):
		self.id = obj.id
		self.state = state
		self.familyName = obj.familyName
		self.Name = symbolName
		self.color = symbolColor
		self.obj = obj
		
def GroupByParameter(lst, func):
	res = {}
	for el in lst:
		key = func(el)
		if key in res:
			res[key].append(el)
		else:
			res[key] = [el]
	return res	
def lists_albums(lists): 
	albums = []
	n = len(lists)
	for i in range(n): 
		j = 0
		NotRepeated = True
		while(j < i): 
			if (lists[j] == lists[i]):
				NotRepeated = False
				break
			j += 1
		if (NotRepeated): 
			albums.append(lists[i])
	return albums	
class PrintSheetsWindow(forms.WPFWindow):
	def __init__(self, xaml_file_name,**kwargs):
		forms.WPFWindow.__init__(self, xaml_file_name)

		self._init_psettings = None
		self._scheduled_sheets = []
		self.sheet_list = kwargs.get('list', None)
		self.list_familyNames.ItemsSource = kwargs.get('symbols', None)
		#self._verify_context()
		self.xx =True
	# doc and schedule
	@property
	def selected_doc(self):
		selected_doc = self.documents_cb.SelectedItem
		for open_doc in revit.docs:
			if open_doc.GetHashCode() == selected_doc.hash:
				return open_doc
	# sheet list
	@property
	def sheet_list(self):
		return self.sheets_lb.ItemsSource

	@sheet_list.setter
	def sheet_list(self, value):
		self.sheets_lb.ItemsSource = value

	@property
	def selected_sheets(self):
		return self.sheets_lb.SelectedItems
	def _verify_context(self):
		new_context = []
		for item in self.sheet_list:
			if not hasattr(item, 'state'):
				new_context.append(BaseCheckBoxItem(item))
			else:
				new_context.append(item)

		self.sheet_list = new_context	
		
	def sheet_selection_changed(self, sender, args):
		#print('List has Changed {}'.format(self.xx))
		self.xx = True
		
	def _set_states(self, state=True,selected=False):
		all_items = self.sheet_list
		if selected:
			current_list = self.selected_sheets
		else:
			current_list = self.sheet_list
		for it in current_list:
			it.state = state
		self.sheet_list = None
		self.sheet_list = all_items

	def check_selected(self, sender, args):
		"""Mark selected checkboxes as checked."""
		#print('Selected {}\n'.format(self.xx))
		if self.xx: 
			self._set_states(state=True, selected=True)
			self.xx = False

	def uncheck_selected(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		if self.xx==True:
			self._set_states(state=False, selected=True)
			self.xx = False

	def button_select(self, sender, args):
		"""Handle select button click."""
		
		try:
			sheets = [x for x in self.sheet_list if x.state]
			# for x in self.sheet_list:
				# print x.state
			selectedSymbol = self.list_familyNames.SelectedItem
			#print selectedSymbol
			if len(sheets)>0:
				if selectedSymbol:
					familyInstances == FilteredElementCollector(doc) \
							.OfClass(FamilyInstance) \
							.ToElements()		
					
					tg = TransactionGroup(doc, "Update")
					tg.Start()
					t=	Transaction(doc, "Calculating")
					t.Start()
					for item in sheets:
						if familyInstance.Id == item.familyInstance_Id:
							my_id = selectedSymbol.id				
							familyInstance.ChangeTypeId(my_id)
							#print item.familySymbol_Id
					self.response={}
					self.Close()
					t.Commit()
					tg.Assimilate()
					MessageBox.Show("Готово!")
				else:
					MessageBox.Show("Выберите типоразмер Основной надписи","Ошибка!")
					#forms.alert('At least one viewport must be selected.')
					
			else:
				MessageBox.Show("Должен быть выбран как минимум один лист","Ошибка!")
		except:
			MessageBox.Show('{}'.format(sys.exc_info()[1]),"Ошибка!")
			#print '{}'.format(NameError)
			self.Close()


#if __name__ == '__main__':
doc = __revit__.ActiveUIDocument.Document
sheets = FilteredElementCollector(doc) \
				.OfClass(ViewSheet) \
				.ToElements()
familyInstances = FilteredElementCollector(doc) \
				.OfClass(FamilyInstance) \
				.OfCategory(BuiltInCategory.OST_TitleBlocks)\
				.ToElements()
familySymbols = FilteredElementCollector(doc) \
				.OfClass(FamilySymbol) \
				.OfCategory(BuiltInCategory.OST_TitleBlocks)\
				.ToElements()

			
list_sheets=[]
#sheets_familyName=[]
list_EnSymbols=[]

# The colors of the states /Enable - Disenable/ for the 
c_En = "#FFA2D88D"
c_DisEn = "#FF1C5507"

list_Symbols= [SymbolOption(x) for x in familySymbols]

list_EnSymbols.append(SymbolStyle(list_Symbols[0],list_Symbols[0].familyName,c_DisEn,False))
#print(list_familySymbols[0].symbolName)
new_familyName = list_Symbols[0].familyName
unsorted_list = []
for el in list_Symbols:
	if el.familyName == new_familyName:	
		unsorted_list.append(el)
	else:
		sorted_list = sorted(unsorted_list, key= lambda SymbolOption:SymbolOption.Name)
		for sy in sorted_list:
			list_EnSymbols.append(SymbolStyle(sy,sy.Name,c_En,True))
		list_EnSymbols.append(SymbolStyle(el,el.familyName,c_DisEn,False))
		# Update SymbolFamily and refill unsortedList el.symbol_Name,c_symEn,True
		unsorted_list = []
		unsorted_list.append(el)
		new_familyName = el.familyName
		
sorted_list = sorted(unsorted_list, key= lambda SymbolOption:SymbolOption.Name)
for sy in sorted_list:
	list_EnSymbols.append(SymbolStyle(sy,sy.Name,c_En,True))
nn=1
for sheet in sheets:
	for familyInstance in familyInstances:
		if familyInstance.OwnerViewId == sheet.Id:
			sheet_name = "{} - {}".format(sheet.SheetNumber,sheet.Name)
			list_name = familyInstance.Name
			list_familyName = familyInstance.Symbol.FamilyName
			format = familyInstance.Symbol.GetParameters("Speech_Формат")
			if format:
				speech_format = format[0].AsString()
			else:
				speech_format = ''
			list_obj = SheetOption(sheet,sheet_name,list_name,list_familyName,
								  familyInstance.Id,familyInstance.Symbol.Id,speech_format,"#FF3CA026",False,'Visible')
			list_sheets.append(list_obj)
			nn+=1

sorted_sheets = []
albums = [x.speech_album for x in list_sheets]
speech_albums = lists_albums(albums)
grouped_sheets = GroupByParameter(list_sheets, func = lambda x: x.speech_album)
for album in speech_albums:
	sorted_sheets.append(SheetOption(sheets[0],album,'','','','','',"#FF1C5507",False,'Hidden'))
	ll = sorted(grouped_sheets[album], key= lambda x:sort_fun(x.Number))
	for z in ll:
		sorted_sheets.append(z)
	#sorted_sheets.append(' ')
#ff = [sort_fun(x.Number) for x in sorted_sheets]
# for x in ff:
	# print x
PrintSheetsWindow('ChangeListsSymbolFamily.xaml',list=sorted_sheets,symbols=list_EnSymbols).ShowDialog()



