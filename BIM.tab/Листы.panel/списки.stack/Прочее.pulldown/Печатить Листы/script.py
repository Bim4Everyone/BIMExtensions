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


import os.path as op
import os
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
							  ScheduleSheetInstance,Element,FilteredElementCollector,ViewSheetSet
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
from pyrevit.forms import pick_folder


class PathWrap:

	def __init__(self):
		self.path = ""
		self.curr = ""
		self.nest = 0

	def set(self, path):
		if path is None: return
		self.path = path 
		self.curr = path

	def get(self): return self.path

	def getparentdirindepth(self, depth):
		if self.path is None or self.path is "" or depth < 0: return ""
		self.nest = 0
		while self.nest < depth: self.curr = self.__getparentdir(self.curr)
		return self.curr

	def getparentdirinlable(self, lable):
		if self.path is None or self.path is "" or lable is None or lable is "": return ""
		find = self.path.find(lable)
		if find < 0: return ""
		self.nest = 0
		count = self.path.count('\\')
		while os.path.basename(self.curr) != lable and self.nest < count: self.curr = self.__getparentdir(self.curr)
		return self.curr

	def __getparentdir(self, path):
		self.nest += 1
		if path is None: return ""
		if os.path.isfile(path): return os.path.dirname(path)
		if os.path.isdir(path): return os.path.dirname(path)
		return ""

Environment.Message("Данная команда временно недоступна!")
Environment.Exit()

# -------------------------------
clslocalpath = os.path.join(os.getenv('APPDATA'), 'pyRevit\extensions\SPEECH_work.extension\SPEECH.tab\ClassLibrary20.dll')
clr.AddReferenceToFileAndPath(clslocalpath)
import PaperSize
from PaperSize import PrintForm


#clr.AddReferenceToFileAndPath(r"W:\Dlls\CustomPrintForm")
#from CustomPrintForm import PrintForm





#pw = PathWrap()
#pw.set(os.path.dirname(os.path.realpath(__file__)))
#issue = pw.getparentdirinlable("ProgramData")

clr.AddReferenceToFileAndPath(os.path.join('C:\Program Files (x86)\IronPython 2.7', "IronPython.dll"))

#--------------------------------------

__title__ = 'Печатить Листы'
logger = script.get_logger()
AvailableDoc = namedtuple('AvailableDoc', ['name', 'hash', 'linked'])
TitleBlockPrintSettings = \
	namedtuple('TitleBlockPrintSettings', ['psettings', 'set_by_param'])
# Non Printable Char
NPC = u'\u200e'
INDEX_FORMAT = '{{:0{digits}}}'

class ViewSheetListItem():
	#def __init__(self, view_sheet,rev_settings=None,state=False): print_settings=None,
	def __init__(self, view_sheet,view_tblock,album,state=False):
		if view_sheet:
			self._sheet = view_sheet
			self._tblock = view_tblock
			self.name = self._sheet.Name
			self.number = self._sheet.SheetNumber
			self.state = state
			self.Visib = 'Visible'
			self.foreGround = "#FF3CA026"
			self.sheet_name = "{} - {}".format(self.number,self.name)
			if view_tblock:
				speechFormat = view_tblock.Symbol.GetParameters("Speech_Формат")
				if speechFormat:
					self.speech_format = speechFormat[0].AsString()
				else:
					self.speech_format=''
			self.speech_album = view_sheet.LookupParameter('Speech_Альбом').AsString()

			self.issue_date = \
				self._sheet.Parameter[
					DB.BuiltInParameter.SHEET_ISSUE_DATE].AsString()
			self.printable = self._sheet.CanBePrinted

			self._print_index = 0
			self._print_filename = "{} - {}".format(self.number,self.name)
			self.group_el = False
		else:
			self._sheet = ''
			self.name = ''
			self.number = ''
			self.state = False
			self.Visib = 'hidden'
			self.sheet_name = album
			self.speech_format=''
			self.foreGround = "#FF1C5507"
			self.printable = False
			self.group_el = True
		#self._print_settings = print_settings
		#self.all_print_settings = print_settings
	@property
	def revit_sheet(self):
		return self._sheet

	@property
	def print_filename(self):
		return self._print_filename

class SheetOption(object):
	def __init__(self, obj,list_name,list_familyName,
					familyInstance_Id,familySymbol_Id,speech_format,state=False):
		self.name = "{} - {}".format(obj.SheetNumber,obj.Name)
		self.familyName = "{} - {}".format(list_name,list_familyName)
		self.sheet_name = obj.Name
		self.id = obj.Id
		self.familyInstance_Id = familyInstance_Id
		self.familySymbol_Id = familySymbol_Id
		self.speech_format = speech_format
		self.obj = obj
		self.revit_sheet = obj
		self.state = state
		self.printable = obj.CanBePrinted
		# def __nonzero__(self):
			# return self.state
		# def __str__(self):
			# return self.name
class SymbolOption(object):
	def __init__(self, obj):
		self.familyName = '{} '.format(obj.FamilyName)
		self.Name = '{} '.format(Element.Name.GetValue(obj))
		self.id = obj.Id
		self.obj = obj
		
class SymbolStyle(object):
	def __init__(self, obj, symbolName,symbolColor, iisEnabled):
		self.id = obj.id
		self.iisEnabled = iisEnabled
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
class PrintSettingListItem(object):
	def __init__(self, print_settings=None):
		self._psettings = print_settings

	@property
	def name(self):
		if isinstance(self._psettings, DB.InSessionPrintSetting):
			return "<In Session>"
		else:
			return self._psettings.Name

	@property
	def print_settings(self):
		return self._psettings

	@property
	def print_params(self):
		if self.print_settings:
			return self.print_settings.PrintParameters

	@property
	def paper_size(self):
		try:
			if self.print_params:
				return self.print_params.PaperSize
		except Exception:
			pass

	@property
	def allows_variable_paper(self):
		return False


class VariablePaperPrintSettingListItem(PrintSettingListItem):
	def __init__(self):
		PrintSettingListItem.__init__(self, None)

	@property
	def name(self):
		return "<Variable Paper Size>"

	@property
	def allows_variable_paper(self):
		return True
		
class PrintSheetsWindow(forms.WPFWindow):
	def __init__(self, xaml_file_name,**kwargs):
		forms.WPFWindow.__init__(self, xaml_file_name)
		#self.doc = __revit__.ActiveUIDocument.Document

		self._init_psettings = None
		self._scheduled_sheets = []
		#self.sheet_list = kwargs.get('list', None)
		#self.list_Printers.ItemsSource = kwargs.get('symbols', None)
		self.sheetlist_set()
		# --------------------------------
		#self._setup_printers()
		#self._setup_print_settings()
		#--------------------------------------
		self._set_Lists()
		# self._setup_docs_list()
		#print(self.selected_doc)
		self.xx =True
		print_mgr = self._get_printmanager()
		#print(print_mgr.ViewSheetSetting.Revert())
	# doc and schedule
	@property
	def selected_doc(self):
		selected_doc = __revit__.ActiveUIDocument.Document
		for open_doc in revit.docs:
			if open_doc.GetHashCode() == selected_doc.GetHashCode():
				#print("Yes")
				return open_doc
	@property
	def sheet_list(self):
		return self.sheets_lb.ItemsSource
	@property
	def revit_sheet(self):
		return self._sheet

	@sheet_list.setter
	def sheet_list(self, value):
		self.sheets_lb.ItemsSource = value
	@property
	def combine_print(self):
		return self.combine_cb.IsChecked
	@property
	def selected_sheets(self):
		return self.sheets_lb.SelectedItems
	@property
	def selected_printer(self):
		return self.printers_cb.SelectedItem
	@property
	def selected_print_setting(self):
		return self.printsettings_cb.SelectedItem
	def _verify_context(self):
		new_context = []
		for item in self.sheet_list:
			if not hasattr(item, 'state'):
				new_context.append(BaseCheckBoxItem(item))
			else:
				new_context.append(item)

		self.sheet_list = new_context
	# def _setup_docs_list(self):
		# if not revit.doc.IsFamilyDocument:
			# docs = [AvailableDoc(name=revit.doc.Title,
								 # hash=revit.doc.GetHashCode(),
								 # linked=False)]
			# docs.extend([
				# AvailableDoc(name=x.Title, hash=x.GetHashCode(), linked=True)
				# for x in revit.query.get_all_linkeddocs(doc=revit.doc)
			# ])
			# self.documents_cb.ItemsSource = docs
			# self.documents_cb.SelectedIndex = 0
	def sheet_selection_changed(self, sender, args):
		#print('List has Changed {}'.format(self.xx))
		self.xx = True
		
	def _set_states(self, state=True,selected=False):
		if self.selected_sheets:
			if self.sheet_list:
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
		self._set_states(state=True, selected=True)
			#self.xx = False

	def uncheck_selected(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self._set_states(state=False, selected=True)
			#self.xx = False
	def edit_formats(self, sender, args):
		editfmt_wnd = \
			EditNamingFormatsWindow(
				'EditNamingFormats.xaml',
				start_with=self.selected_naming_format
				)
		editfmt_wnd.show_dialog()
		self.namingformat_cb.ItemsSource = editfmt_wnd.naming_formats
		self.namingformat_cb.SelectedItem = editfmt_wnd.selected_naming_format
# ---------------------------------------------------------------------------------  Setup PrintSetting
	def _get_printmanager(self):
		try:
			return self.selected_doc.PrintManager
		except Exception as printerr:
			logger.critical('Error getting printer manager from document. '
							'Most probably there is not a printer defined '
							'on your system. | %s', printerr)
			return None
			
	def _setup_printers(self):
		printers = list(Drawing.Printing.PrinterSettings.InstalledPrinters)
		self.printers_cb.ItemsSource = printers
		print_mgr = self._get_printmanager()
		self.printers_cb.SelectedItem = print_mgr.PrinterName

	def get_print_settings(self):
		print_settings = [VariablePaperPrintSettingListItem()]
		print_settings.extend(
			[PrintSettingListItem(self.selected_doc.GetElement(x))
			 for x in self.selected_doc.GetPrintSettingIds()]
			)
		print_mgr = self._get_printmanager()
		if isinstance(print_mgr.PrintSetup.CurrentPrintSetting,
					  DB.InSessionPrintSetting):
			in_session = PrintSettingListItem(
				print_mgr.PrintSetup.CurrentPrintSetting
				)
			print_settings.append(in_session)

		return print_settings
	def _setup_print_settings(self):
		print_settings = [VariablePaperPrintSettingListItem()]
		print_settings.extend(
			[PrintSettingListItem(self.selected_doc.GetElement(x))
			 for x in self.selected_doc.GetPrintSettingIds()]
			)
		print_mgr = self._get_printmanager()
		if isinstance(print_mgr.PrintSetup.CurrentPrintSetting,
					  DB.InSessionPrintSetting):
			in_session = PrintSettingListItem(
				print_mgr.PrintSetup.CurrentPrintSetting
				)
			print_settings.append(in_session)
			self.printsettings_cb.SelectedItem = in_session
		else:
			self._init_psettings = print_mgr.PrintSetup.CurrentPrintSetting
			cur_psetting_name = print_mgr.PrintSetup.CurrentPrintSetting.Name
			for psetting in print_settings:
				if psetting.name == cur_psetting_name:
					self.printsettings_cb.SelectedItem = psetting
		# for x in print_settings:
			# print '{} ---- {}'.format(x.name,x.print_params)
		self.printsettings_cb.ItemsSource = print_settings	
		return print_settings
# -------------------------------------------------------------------------------------- Set Sheets  
	def _get_ordered_schedule_sheets(self):
		sheets = DB.FilteredElementCollector(self.selected_doc)\
				.OfClass(ViewSheet)\
				.WhereElementIsNotElementType()\
				.ToElements()
		return sheets
	def _find_sheet_tblock(self, revit_sheet, tblocks):
		for tblock in tblocks:
			view_sheet = revit_sheet.Document.GetElement(tblock.OwnerViewId)
			if view_sheet.Id == revit_sheet.Id:
				return tblock
		return sheets
	
	def _lists_albums(self,set_lists): 
		albums = []
		n = len(set_lists)
		for i in range(n): 
			j = 0
			NotRepeated = True
			while(j < i): 
				if (set_lists[j] == set_lists[i]):
					NotRepeated = False
					break
				j += 1
			if (NotRepeated): 
				albums.append(set_lists[i])
		return albums
	def sort_fun(self,str):
		list = [int(x) for x in re.findall(r'\b\d+\b', str)]
		return list[0]	
	def sheetlist_set(self):
		tblocks = FilteredElementCollector(self.selected_doc) \
		.OfClass(FamilyInstance) \
		.OfCategory(BuiltInCategory.OST_TitleBlocks)\
		.ToElements()
		#sheet_printsettings = self._get_sheet_printsettings(tblocks,self.printsettings_cb.ItemsSource)
		self._scheduled_sheets = [
			ViewSheetListItem(
				view_sheet=x,
				view_tblock=self._find_sheet_tblock(x, tblocks),
				album='',
				state=False)
			for x in self._get_ordered_schedule_sheets()
			]
		sorted_sheets = []
		albums = [x.speech_album for x in self._scheduled_sheets]
		speech_albums = self._lists_albums(albums)
		grouped_sheets = GroupByParameter(self._scheduled_sheets, func = lambda x: x.speech_album)
		for album in speech_albums:
			#print self._get_ordered_schedule_sheets()[0].Name
			sorted_sheets.append(ViewSheetListItem('','',album,False))
			ll = sorted(grouped_sheets[album], key= lambda x:self.sort_fun(x.number))
			for z in ll:
				sorted_sheets.append(z)
		
		self.sheet_list = sorted_sheets
#---------------------------------------------------------------------------------------------- Print Functions
	def check_formats(self):
		# get already printsettings that we have
		print_mgr = self._get_printmanager()
		paper_formats = [paper.Name for paper in print_mgr.PaperSizes]
		printsettings = [x.name for x in list(self.get_print_settings())]
		forms_0_hasFound = 0
		forms_1_hasFound = 0
		forms_0_hasAdded = 0
		forms_1_hasAdded = 0

		for p_format in printsettings:
			#print p_format
			[p_size1,ori1,check_for1] = self.check_form1(p_format)
			if check_for1:
				if p_size1 in paper_formats:
					forms_0_hasFound += 1
					#break
				else:
					self.add_newForm1(paper_size)
					forms_0_hasAdded += 1

			[height_form,width_form,ori2,check_for2] = self.check_form2(p_format)
			if check_for2:
				p_size2 =  '{:.0f}x{:.0f}'.format(height_form,width_form)
				if p_size2 in paper_formats:
					forms_1_hasFound += 1
					#break
				else:
					self.add_newForm2(p_size2,height_form,width_form)
					forms_1_hasAdded += 1

		# print 'form0 found: {} Added: {}'.format(forms_0_hasFound,forms_0_hasAdded)
		# print 'form1 found: {} Added: {}'.format(forms_1_hasFound,forms_1_hasAdded)
		# print 'OF {}'.format(len(printsettings))


	def print_sheets(self, sender, args):
		# Check if we have paper formats we have to add
		self.check_formats()
		if self.sheet_list:
			target_sheets =[x for x in self.sheet_list if x.state and not x.group_el]
			# confirm print if a lot of sheets are going to be printed
			printable_count = len([x for x in target_sheets if x.printable and not x.group_el])
			if printable_count > 5:
				# prepare warning message
				sheet_count = len(target_sheets)
				message = str(printable_count)
				if printable_count != sheet_count:
					message += ' (out of {} total)'.format(sheet_count)
#ok=True, yes=False, no=False
				if not forms.alert('Are you sure you want to print {} '
									'sheets individually? The process can '
									'not be cancelled.'.format(message),no=True,yes=True):
					return
			self._print_sheets_in_order(target_sheets)


	def check_form2(self,st):
		w = 0
		h = 0
		check = False
		orientation = ''

		sep = ['','']
		# sperator0 in English
		if 'K' in st:
			sep[1] = 'K'
			orientation = 'K' # in en
		elif 'A' in st:
			sep[1] = 'A'
			orientation = 'A' # in en
		# sperator0 in Russian
		elif 'К' in st:
			sep[1] = 'К'
			orientation = 'K' # in en
		elif 'А' in st:
			sep[1] = 'А'
			orientation = 'A' # in en
		# ------------------
		if 'x' in st:
			sep[0] = 'x'
		elif 'х' in st:
			sep[0] = 'х'

		if sep[1]=='K' or sep[1]=='К' or sep[1]=='A' or sep[1]=='А':
			if sep[0]=='x' or sep[0]=='х':
				s1 = st.split(sep[1])
				s0 = s1[0].split(sep[0])
				#math.
				# Add the new Form
				h = float(s0[0])
				w = float(s0[1])
				#orientation = sep[1]
				if isinstance(h,float) and isinstance(w,float):
					check = True
		
		return[h,w,orientation,check]

	def check_form1(self,st):
		ori_array_ru = ['А','К']
		ori_array_en = ['A','K']

		paper_array_ru = ['А0','А1','А2','А3','А4']
		paper_array_en = ['A0','A1','A2','A3','A4']
		# initialize values
		check_for = False
		paper_size = ''
		orientation = ''

		p_size = st[0:2]
		ori = st[2]
		#print '{} - {}'.format(p_size,ori)
		if ori in ori_array_en:
			#print 'ori in ori_array_en'
			if p_size in paper_array_en:
				#print 'p_size in paper_array_en'
				check_for = True
				paper_size = p_size
				orientation = ori
				return[paper_size,orientation,check_for]
			if p_size in paper_array_ru:
				#print 'p_size in paper_array_ru'
				ind_paper = paper_array_ru.index(p_size)
				check_for = True
				paper_size = paper_array_en[ind_paper]
				orientation =ori
				return[paper_size,orientation,check_for]
		if ori in ori_array_ru:
			ind_ori = ori_array_ru.index(ori)
			#print 'ori in ori_array_ru'
			if p_size in paper_array_ru:
				#print 'p_size in paper_array_ru'
				ind_paper = paper_array_ru.index(p_size)
				check_for = True
				paper_size = paper_array_en[ind_paper]
				orientation = ori_array_en[ind_ori]
				return[paper_size,orientation,check_for]
			if p_size in paper_array_en:
				#print 'p_size in paper_array_en'
				check_for = True
				paper_size = p_size
				orientation = ori_array_en[ind_ori]
				return[paper_size,orientation,check_for]

		return[paper_size,orientation,check_for]

	def add_newForm1(self,paper_size):
		if paper_size =='A0' or paper_size =='А0':
			self.add_newForm2(paper_size,841,1189)
		elif paper_size =='A1' or paper_size =='А1':
			self.add_newForm2(paper_size,594,841)
		elif paper_size =='A2' or paper_size =='А2':
			self.add_newForm2(paper_size,420,594)
		elif paper_size =='A3' or paper_size =='А3':
			self.add_newForm2(paper_size,297,420)
		elif paper_size =='A4' or paper_size =='А4':
			self.add_newForm2(paper_size,210,297)


	def add_newForm2(self,name_form,width_form,height_form):
		PrintForm.AddCustomPaperSize(name_form,width_form,height_form)

	def str2en(seld,st):
		ru_array = ['А','К','х']
		en_array = ['A','K','x']
		st_en = list(st)
		for ch in st_en:
			if ch in ru_array:
				ind0 = st_en.index(ch)
				ind1 = ru_array.index(ch)
				st_en[ind0] = en_array[ind1]
		return ''.join(st_en)



	def _print_sheets_in_order(self, target_sheets):
		# make sure we can access the print config
		print_mgr = self._get_printmanager()
		print_mgr.PrintToFile = True
		if not print_mgr:
			return
		print_mgr.SelectNewPrintDriver('PDFCreator')
		print_mgr.PrintRange = DB.PrintRange.Current

		#per_sheet_psettings = self.selected_print_setting.allows_variable_paper
		with revit.TransactionGroup('Set Printer Settings',
								  doc=self.selected_doc):
			# Collect existing sheet sets
			for sheet in target_sheets:
				# get already printsettings that we have
				printsettings = list(self.get_print_settings())
				printsettings_names = [x.name for x in printsettings]
				orientation = ''
				 
				Print_setup = self.str2en(sheet.speech_format)
				#print Print_setup
				#Print_setup = sheet.speech_format
				if sheet.printable:
					#try:
					#print sheet.speech_format
					with revit.Transaction('Add Printer setup',doc=self.selected_doc):
						Print_setup_hasfound = False
						if Print_setup in printsettings_names:
							ind = printsettings_names.index(Print_setup)
							#print '{} -{}'.format(ind,printsettings[ind-1])
							print_mgr.PrintSetup.CurrentPrintSetting = printsettings[ind].print_settings
							#print print_mgr.PrintSetup.CurrentPrintSetting.Name
							Print_setup_hasfound = True
							#print 'print setup has found'
					if not Print_setup_hasfound:
						settings = print_mgr.PrintSetup.InSession
						paper_size_hasfound = False
						# get the string
						st = self.str2en(sheet.speech_format)
						if len(st) == 3:
							#print 'paper_size first format'
							[p_size,ori1,check_for1] = self.check_form1(st)
							#print '{}-{} {}'.format(p_size,ori1,check_for1)
							if not check_for1:
								RightFormWindow.show([],'Неправильный Speech_Формат',sheet=sheet)
								continue
							paper_size = p_size
							orientation = ori1
							Print_setup = '{}{}'.format(p_size,ori1)
							#print 'paper_size first format has checked'
							# Set PaperSize
							# --> check if paper size already have
							for size in print_mgr.PaperSizes:
								if size.Name ==  paper_size:
									settings.PrintParameters.PaperSize = size
									paper_size_hasfound = True
									#print 'paper_size first format has found'
							# if Not, Add it
							if not paper_size_hasfound:
								with revit.Transaction('Add Paper Size',doc=self.selected_doc):
									self.add_newForm1(paper_size)
									#print 'paper_size first format has added'
						else:
							#print 'paper_size second format'
							[h,w,ori2,check_for2] = self.check_form2(st)
							if not check_for2:
								RightFormWindow.show([],'Неправильный Speech_Формат',sheet=sheet)
								continue
							width_form = min([w,h])
							height_form = max([w,h])
							#paper_size = st[0:len(st)-1]
							paper_size =  '{:.0f}x{:.0f}'.format(h,w)
							orientation = ori2
							Print_setup = '{}{}'.format(paper_size,ori2)
							# Set PaperSize
							# --> check if paper size already have
							for size in print_mgr.PaperSizes:
								if size.Name ==  sheet.speech_format:
									settings.PrintParameters.PaperSize = size
									paper_size_hasfound = True
									#print 'paper_size second format has found'
							# if Not, Add it
							if not paper_size_hasfound:
								with revit.Transaction('Add Paper Size',doc=self.selected_doc):
									self.add_newForm2(paper_size,width_form,height_form)
									#print 'paper_size second format has added'
						print_mgr = self._get_printmanager()
						print_mgr.SelectNewPrintDriver('PDFCreator')
						print_mgr.PrintRange = DB.PrintRange.Current
						settings = print_mgr.PrintSetup.InSession
						for size in print_mgr.PaperSizes:
							#print size.Name
							if size.Name ==  paper_size:
								settings.PrintParameters.PaperSize = size
								#print 'real paper_size has found'

							
						# choose paper orientation Landscape or Portrait Landscape
						settings.PrintParameters.PageOrientation = print_mgr.PrintSetup.CurrentPrintSetting\
														.PrintParameters.PageOrientation.Landscape
						if orientation=='K': 
							settings.PrintParameters.PageOrientation = print_mgr.PrintSetup.CurrentPrintSetting\
																				.PrintParameters.PageOrientation.Portrait		
						# Choose ZoomType Zoom or FitToPage
						# settings.PrintParameters.ZoomType = print_mgr.PrintSetup.CurrentPrintSetting\
						# 										.PrintParameters.ZoomType.FitToPage
						settings.PrintParameters.ZoomType = print_mgr.PrintSetup.CurrentPrintSetting\
																	.PrintParameters.ZoomType.Zoom
						settings.PrintParameters.Zoom = 100
						settings.PrintParameters.PaperPlacement = print_mgr.PrintSetup.CurrentPrintSetting\
																			.PrintParameters.PaperPlacement.Center

						print_mgr.PrintSetup.CurrentPrintSetting = settings
					print_fileName = sheet.sheet_name+'.pdf'
					print_mgr.PrintToFile = True
					print_mgr.PrintToFileName = op.join(r'C:\\',print_fileName)
					with revit.Transaction('Save Printer setup',
						doc=self.selected_doc):
						#print_mgr.PrintSetup.CurrentPrintSetting = sheet.print_settings
						if  not Print_setup_hasfound:
							print_mgr.PrintSetup.SaveAs(Print_setup)
							#print 'print setup has added'
						print_mgr.Apply()
						print_mgr.SubmitPrint(sheet.revit_sheet)
					
					# except:
					# 	MessageBox.Show('{}'.format(sys.exc_info()[1]),"Ошибка!")
					# 	#MessageBox.Show('','Error')

				else:
					logger.debug('Sheet %s is not printable. Skipping print.',
								 sheet.number)

	# ------------------------------------------------------------------- Work with view sets ----------------
	def save_set(self, sender, args):
		# make sure we can access the print config
		print_mgr = self._get_printmanager()
		sheet_set = DB.ViewSet()
		set_lists = [li for li in self.sheet_list if li.state]
		for li in set_lists:
			if not li.group_el:
				rvtsheet = li.revit_sheet
				sheet_set.Insert(rvtsheet)
		with revit.TransactionGroup('Add New ViewSet',
										doc=self.selected_doc):
			if not print_mgr:
				return 
			with revit.Transaction('Get CurrentPrintSetting',
								doc=self.selected_doc):
				# print_mgr.PrintSetup.CurrentPrintSetting = \
				# 	self.selected_print_setting.print_settings
				print_mgr.SelectNewPrintDriver('PDFCreator')
				print_mgr.PrintRange = DB.PrintRange.Select
			with revit.Transaction('Remove ViewSheet Set',
									doc=self.selected_doc):
				# Delete existing matching sheet set
				current_ViewSet = self.set_Lists.SelectedItem
				name = current_ViewSet.Name
				print_mgr.ViewSheetSetting.CurrentViewSheetSet = \
					self.set_Lists.SelectedItem
				print_mgr.ViewSheetSetting.Delete()
				# print_mgr.ViewSheetSetting.CurrentViewSheetSet.Views =sheet_set
				# print print_mgr.ViewSheetSetting.CurrentViewSheetSet.Name
				# print_mgr.ViewSheetSetting.Save()
		self.saveSet_as('','',sheetsetname=name,sheetSet = sheet_set)
			#self._set_Lists(0)
	def saveSet_as(self,sender,args,**kwargs):
		sheetsetname =  kwargs.get('sheetsetname')
		print_mgr = self._get_printmanager()
		sheet_set = kwargs.get('sheetSet')
		#index = kwargs.get('index')
		if not sheet_set:
			sheet_set = DB.ViewSet()
			set_lists = [li for li in self.sheet_list if li.state]
			for li in set_lists:
				if not li.group_el:
					rvtsheet = li.revit_sheet
					sheet_set.Insert(rvtsheet)
		print_mgr = self._get_printmanager()
		if not sheetsetname:
			sheetsetname = EnterSetViewName.show([],'Введите Название SetView',button_ok='Ok',button_close='Отмена',field_name="ИМЯ")
		#print sheetsetname
		if 	sheetsetname:
			with revit.TransactionGroup('Add New ViewSet',
										doc=self.selected_doc):
				if not print_mgr:
					return ##print_mgr.PrintSetup.InSession
				with revit.Transaction('Get CurrentPrintSetting',
									doc=self.selected_doc):
					# print_mgr.PrintSetup.CurrentPrintSetting = \
					# 	self.selected_print_setting.print_settings
					print_mgr.SelectNewPrintDriver('PDFCreator')
					print_mgr.PrintRange = DB.PrintRange.Select


				with revit.Transaction('Save the new ViewSet',
									doc=self.selected_doc):
					try:
						viewsheet_settings = print_mgr.ViewSheetSetting
						viewsheet_settings.CurrentViewSheetSet.Views = sheet_set
						viewsheet_settings.SaveAs(sheetsetname)
						
					except Exception as viewset_err:
						logger.critical(
							'Error setting sheet set on print mechanism. '
							'These items are included in the viewset '
							'object:\n')
						raise viewset_err
				viewSheetSets = self.get_viewSheetSets()
				#if not index:
				index = len(viewSheetSets) - 1 
				self._set_Lists(index)

	def rename_set(self, sender, args):
		print_mgr = self._get_printmanager()
		viewSheetSets = self.get_viewSheetSets()
		sets_names = [s.Name for s in viewSheetSets]
		index = sets_names.index(self.set_Lists.SelectedItem.Name)

		sheetsetname = EnterSetViewName.show([],'Введите Название SetView',\
					            button_ok='Ok',button_close='Отмена',field_name="ИМЯ")
		if sheetsetname:
			if sheetsetname in sets_names:
				MessageBox.Show("The set already exist","Error")
				#raise SystemExit(1)
			else:
				with revit.TransactionGroup('Add New ViewSet',
												doc=self.selected_doc):
					if not print_mgr:
						return ##print_mgr.PrintSetup.InSession
					with revit.Transaction('Get CurrentPrintSetting',
										doc=self.selected_doc):
						# print_mgr.PrintSetup.CurrentPrintSetting = \
						# 	self.selected_print_setting.print_settings
						print_mgr.SelectNewPrintDriver('PDFCreator')
						print_mgr.PrintRange = DB.PrintRange.Select
					with revit.Transaction('Remove ViewSheet Set',
											doc=self.selected_doc):
						# Delete existing matching sheet set
						current_ViewSet = self.set_Lists.SelectedItem
						name = current_ViewSet.Name
						print_mgr.ViewSheetSetting.CurrentViewSheetSet = \
							self.set_Lists.SelectedItem
						print_mgr.ViewSheetSetting.Rename(sheetsetname)
						self._set_Lists(index)

	def delete_set(self, sender, args):
		print_mgr = self._get_printmanager()
		viewSheetSets = self.get_viewSheetSets()
		sets_names = [s.Name for s in viewSheetSets]
		index = sets_names.index(self.set_Lists.SelectedItem.Name)
		with revit.TransactionGroup('Add New ViewSet',
										doc=self.selected_doc):
			if not print_mgr:
				return ##print_mgr.PrintSetup.InSession
			with revit.Transaction('Get CurrentPrintSetting',
								doc=self.selected_doc):
				# print_mgr.PrintSetup.CurrentPrintSetting = \
				# 	self.selected_print_setting.print_settings
				print_mgr.SelectNewPrintDriver('PDFCreator')
				print_mgr.PrintRange = DB.PrintRange.Select
			with revit.Transaction('Remove ViewSheet Set',
									doc=self.selected_doc):
				# Delete existing matching sheet set
				print_mgr.ViewSheetSetting.CurrentViewSheetSet = \
					self.set_Lists.SelectedItem
				print_mgr.ViewSheetSetting.Delete()
			if index:
				self._set_Lists(index-1)
			else:
				self._set_Lists(index)
#---------------------------------------------------------------------------------- Set View Sets ----------------			
	def get_viewSheetSets(self):
		lists = FilteredElementCollector(doc).OfClass(ViewSheetSet)\
				.WhereElementIsNotElementType()\
				.ToElements()
		return lists
	def _set_Lists(self,index=0):
		viewSheetSets = self.get_viewSheetSets()
		size_sets =  len(viewSheetSets)
		if size_sets:
			self.set_Lists.SelectedItem = viewSheetSets[index]
			self.set_Lists.ItemsSource = viewSheetSets
		
	def	set_lists_changed(self,sender, args):
		#print 'dfgd'
		all_lists = self.sheet_list
		set_List = self.set_Lists.SelectedItem
		#print '1: {}'.format(set_List.Name)
		set_views = set_List.Views
		for lis in all_lists:
			lis.state = False
		for lis in all_lists:
			for v in set_views:
				if lis.name == v.Name:
					lis.state = True
					#print v.Name
		# for l in all_lists:
		# 	print l.name
		self.sheets_lb.ItemsSource = None
		self.sheets_lb.ItemsSource = all_lists
class EnterSetViewName(forms.TemplateUserInputWindow):
	xaml_source = op.join(op.dirname(__file__),'BaseWindow.xaml')
	
	def _setup(self, **kwargs):		
		button_ok = kwargs.get('button_ok', None)
		if button_ok:
			self.addField.Content = button_ok
		button_close = kwargs.get('button_close', None)
		if button_close:
			self.closeField.Content = button_close
		field_name = kwargs.get('field_name', None)
		if field_name:
			self.setfield.Text = field_name

	def insert_Name(self, sender, args):
		"""Handle select button click."""
		self.response = self.NameField.Text
		self.Close()
	def close_Field(self, sender, args):
		self.Close()
class RightFormWindow(forms.TemplateUserInputWindow):
	xaml_source = op.join(op.dirname(__file__),'FormWindow.xaml')
	
	def _setup(self, **kwargs):		
	 	sheet = kwargs.get('sheet', None)
	 	self.sheetName.Text = sheet.sheet_name
		#self.sheetName = self._context

	# def _list_options(self, checkbox_filter=None):
	# 	self.sheetName = self._context
	def close_Field(self, sender, args):
		self.Close()
# ---------------------------------------------------------	
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
			list_name = familyInstance.Name
			list_familyName = familyInstance.Symbol.FamilyName
			format = familyInstance.Symbol.GetParameters("Speech_Формат")
			if format:
				speech_format = format[0].AsString()
			else:
				speech_format = ''
			list_obj = SheetOption(sheet,list_name,list_familyName,
								  familyInstance.Id,familyInstance.Symbol.Id,speech_format)
			list_sheets.append(list_obj)
			nn+=1
#for x in speech_format:
#	print x.AsString
window = PrintSheetsWindow('PrintSheets.xaml',list=list_sheets,symbols=list_EnSymbols)
#try:
window.ShowDialog()
# except :
# 	#traceback.print_exc(file=sys.stdout)
# 	MessageBox.Show('{}'.format(sys.exc_info()[1]),"Ошибка!")
# 	window.Close()
# 	#sys.exit(1)




