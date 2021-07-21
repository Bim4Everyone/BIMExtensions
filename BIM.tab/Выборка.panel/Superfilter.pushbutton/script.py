# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
import clr
import math
import time
import collections
from shutil import copyfile
from pyrevit import forms
from pyrevit import revit

from pyrevit.framework import Controls
clr.AddReference('System')
clr.AddReference('System.IO')
clr.AddReference("System.Windows.Forms")
clr.AddReference("PresentationCore")

import System.Windows.Media
clr.AddReference('WindowsBase')
clr.AddReference('PresentationFramework')
from System.Windows.Media import VisualTreeHelper

from System.Windows.Data import CollectionViewSource,PropertyGroupDescription
from System.IO import FileInfo
from System.Windows.Forms import MessageBox, SaveFileDialog, DialogResult
from System.Collections.Generic import List
from Autodesk.Revit.DB import DetailNurbSpline, CurveElement, ElementTransformUtils, \
							  DetailLine, View, ViewDuplicateOption, XYZ, LocationPoint, \
							  TransactionGroup, Transaction, FilteredElementCollector, \
							  ElementId, BuiltInCategory, FamilyInstance, ViewDuplicateOption, \
							  ViewSheet, FamilySymbol, Viewport, DetailEllipse, DetailArc, TextNote, \
							  ScheduleSheetInstance,Element,FilteredElementCollector,ViewSheetSet,SelectionFilterElement,BuiltInParameter
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
from pyrevit import script
import pyrevit

__title__ = 'Суперфильтр'
__doc__ = 'Выбор элементов в проекте на основе значений параметров'

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

class QuickSelectWindow(forms.WPFWindow):
	def __init__(self, xaml_file_name,**kwargs):
		forms.WPFWindow.__init__(self, xaml_file_name)
		self.cb_current_view.IsEnabled = __revit__.ActiveUIDocument.Document.ActiveView != None
        
		self._init_psettings = None        
		cats = kwargs.get('list', None)        
		self.elments = kwargs.get('Elements', None)        
		self.list_cats.ItemsSource = cats

	def fun_cat(self,cat_checked,cat_status):

		all_cats = self.list_cats.ItemsSource 
		selected_cats = self.list_cats.SelectedItems
		# save the expanded categories
		i = 0
		expanded_cats = []
		cats = self.list_cats
		if len(cats.SelectedItems) > 0:
			for cat in cats.SelectedItems:
				cat.cat_state = cat_status

		for cat in cats.ItemsSource:
			container = cats.ItemContainerGenerator.ContainerFromItem(cat)
			child_0 = VisualTreeHelper.GetChild(container,0)
			child_1 = VisualTreeHelper.GetChild(child_0,0)
			dataTemplate = child_1.ContentTemplate
			i += 1
			cat_checkBox = dataTemplate.FindName("cat_checkBox",child_1)
			if cat_checkBox.IsChecked == cat_checked:
				cat.cat_state = cat_status

		pos_scroller = self.myScroll.VerticalOffset
		self.list_cats.ItemsSource = []
		self.list_cats.ItemsSource = all_cats
		self.myScroll.ScrollToVerticalOffset(pos_scroller)

	def selected_cat(self, sender, args):
		"""Mark selected categories as checked."""
		self.fun_cat(True,True)

	def unselected_cat(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self.fun_cat(False,False)

	def indeter_cat(self,sender,args):
		"""Mark intermediate checkboxes as unchecked."""
		self.fun_cat(None,False)

	def MyMouseWheelV(self,sender,args):
		self.myScroll.ScrollToVerticalOffset(self.myScroll.VerticalOffset - args.Delta)

	def select_els(self,sender,args):
		"""Mark selected checkboxes as unchecked."""
		# selected elements
		returned_els = []
		all_cats = self.list_cats.ItemsSource 
		cats = self.list_cats
		selected_cats = [] 
		for cat in cats.ItemsSource:
			if cat.cat_state == True:
				selected_cats.append(cat.cat_Name)
		for el in self.elments:
			if el.Category.Name in selected_cats:
				returned_els.append(el.Id)

		if self.cb_current_view.IsChecked:
			document = __revit__.ActiveUIDocument.Document
			activeViewId = __revit__.ActiveUIDocument.Document.ActiveView.Id
			activeViewElements = FilteredElementCollector(document, activeViewId).ToElements()
			activeViewElements = [ element for element in activeViewElements if element.Parameter[BuiltInParameter.WALL_HEIGHT_TYPE] != None and element.Parameter[BuiltInParameter.WALL_HEIGHT_TYPE].AsElementId() == ElementId.InvalidElementId ]
			
			activeViewElementsIds = [ element.Id for element in activeViewElements ]
			returned_els = [x for x in returned_els if x in activeViewElementsIds]
				
		if len(returned_els) > 0:
			__revit__.ActiveUIDocument.Selection.SetElementIds(List[DB.ElementId](returned_els))
			self.Close()
		else:
			#forms.alert('Выделенных элементов нет !!')
			MessageBox.Show("Выделенных элементов нет !!","Внимание!")

		

	def select_cats(self,sender,args):
		all_cats = self.list_cats.ItemsSource 
		cats = self.list_cats
		selected_cats = [] 
		for cat in cats.ItemsSource:
			if cat.cat_state == True:
				selected_cats.append(cat.cat_Name)
		list_cats = [category(cat,self.elments) for cat in selected_cats]
		if len(selected_cats) > 0:
			self.Close()
			res = PrintSheetsWindow('quickSelect.xaml',list = list_cats).ShowDialog()
			
			if not res:
				output = pyrevit.output.get_output()
				output.close()
		else:
			#forms.alert('ОШИБКА!!...  Выделите нужные категории')
			MessageBox.Show("Выделите нужные категории !!","ОШИБКА!")

	def cancel(self,sender,args):
		"""Mark selected checkboxes as unchecked."""
		self.Close()

class PrintSheetsWindow(forms.WPFWindow):
	def __init__(self, xaml_file_name,**kwargs):

		forms.WPFWindow.__init__(self, xaml_file_name)
		self._init_psettings = None
		cats = kwargs.get('list', None)
		self.list_cats.ItemsSource = cats
		self.and_checked = False
	@property
	def selected_doc(self):
		selected_doc = self.documents_cb.SelectedItem
		for open_doc in revit.docs:
			if open_doc.GetHashCode() == selected_doc.hash:
				return open_doc
	# sheet list
	@property
	def sheet_list(self):
		return self.list_cats.ItemsSource

	@sheet_list.setter
	def sheet_list(self, value):
		self.list_cats.ItemsSource = value

	@property
	def selected_sheets(self):
		return self.list_cats.SelectedItems

	def _set_states(self, state=True,selected=False):
		if self.selected_sheets:
			if self.sheet_list:
				all_items = self.sheet_list
				current_list = self.selected_sheets
				for it in current_list:
					it.state = state

					
				self.sheet_list = []
				self.sheet_list = all_items
	def fun_cat(self,cat_checked,cat_status):
		all_cats = self.list_cats.ItemsSource 
		selected_cats = self.list_cats.SelectedItems
		# save the expanded categories
		i = 0
		expanded_cats = []
		cats = self.list_cats
		if len(cats.SelectedItems) > 0:
			for cat in cats.SelectedItems:
				cat.cat_state = cat_status
				for par in cat.paramters:
					par.par_state = cat_status
					for el in par.elements:
						el.el_state = cat_status

		for cat in cats.ItemsSource:
			container = cats.ItemContainerGenerator.ContainerFromItem(cat)
			child_0 = VisualTreeHelper.GetChild(container,0)
			child_1 = VisualTreeHelper.GetChild(child_0,0)
			dataTemplate = child_1.ContentTemplate
			cat_exp = dataTemplate.FindName("cat_expander",child_1)
			if cat_exp.IsExpanded == True:
				expanded_cats.append(i)
			i += 1
			cat_checkBox = dataTemplate.FindName("cat_checkBox",child_1)
			if cat_checkBox.IsChecked == cat_checked:
				cat.cat_state = cat_status
				for par in cat.paramters:
					par.par_state = cat_status
					for el in par.elements:
						el.el_state = cat_status

		pos_scroller = self.myScroll.VerticalOffset
		self.list_cats.ItemsSource = []
		self.list_cats.ItemsSource = all_cats
		# reset the expanded categories
		cats = self.list_cats
		i = 0 
		#
		self.list_cats.UpdateLayout()
		for cat in cats.ItemsSource:
			container = cats.ItemContainerGenerator.ContainerFromItem(cat)
			child_0 = VisualTreeHelper.GetChild(container,0)
			child_1 = VisualTreeHelper.GetChild(child_0,0)
			dataTemplate = child_1.ContentTemplate
			cat_exp = dataTemplate.FindName("cat_expander",child_1)
			if i in expanded_cats:
				cat_exp.IsExpanded = True
			i += 1
		self.myScroll.ScrollToVerticalOffset(pos_scroller)

	def selected_cat(self, sender, args):
		"""Mark selected categories as checked."""
		self.fun_cat(True,True)

	def unselected_cat(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self.fun_cat(False,False)

	def indeter_cat(self,sender,args):
		"""Mark intermediate checkboxes as unchecked."""
		self.fun_cat(None,False)

	def fun_par(self,checked):
		"""Mark selected checkboxes as unchecked."""
		# save the expanded categories
		i = 0
		expanded_cats = []
		expanded_pars = []

		# set the new states
		all_cats = self.list_cats.ItemsSource 
		cats = self.list_cats
		self.list_cats.UpdateLayout()
		for cat in cats.ItemsSource:
			cat_container = cats.ItemContainerGenerator.ContainerFromItem(cat)
			child_0 = VisualTreeHelper.GetChild(cat_container,0)
			child_1 = VisualTreeHelper.GetChild(child_0,0)
			dataTemplate = child_1.ContentTemplate
			cat_exp = dataTemplate.FindName("cat_expander",child_1)
			if cat_exp.IsExpanded == True:
				list_pars = dataTemplate.FindName("list_pars",child_1)
				expanded_cats.append(i)
				j = 0
				checked_pars = 0
				none_pars = 0
				cat.cat_state = False
				if len(list_pars.SelectedItems) > 0:
					for par in list_pars.SelectedItems:
						par.par_state = checked
						for el in par.elements:
							el.el_state = checked
					for par in list_pars.ItemsSource:
						if par.par_state == True:
							checked_pars += 1
				else:
					list_pars.UpdateLayout()
					for par in list_pars.ItemsSource:
						par_container = list_pars.ItemContainerGenerator.ContainerFromItem(par)
						child_0 = VisualTreeHelper.GetChild(par_container,0)
						child_1 = VisualTreeHelper.GetChild(child_0,0)
						dataTemplate = child_1.ContentTemplate
						par_checkBox = dataTemplate.FindName("par_checkBox",child_1)
						par_exp = dataTemplate.FindName("par_expander",child_1)
						if par_exp.IsExpanded == True:
							expanded_pars.append(i * 10 + j)
						if par_checkBox.IsChecked == True:
							checked_pars += 1
							par.par_state = True
							for el in par.elements:
								el.el_state = True
						if par_checkBox.IsChecked == False:
							par.par_state = False
							for el in par.elements:
								el.el_state = False
						if par_checkBox.IsChecked == None:
							if not par.par_state == None:
								par.par_state = not par.par_state
								for el in par.elements:
									el.el_state = not el.el_state
							else:
								none_pars += 1
						j += 1

				
				if checked_pars == len(list_pars.ItemsSource):
					cat.cat_state = True
				elif checked_pars > 0 or none_pars > 0:
					cat.cat_state = None

			i += 1
		pos_scroller = self.myScroll.VerticalOffset
		self.list_cats.ItemsSource = []
		self.list_cats.ItemsSource = all_cats
		# reset the expanded categories
		all_cats = self.list_cats.ItemsSource 
		cats = self.list_cats
		i = 0
		self.list_cats.UpdateLayout()
		for cat in cats.ItemsSource:
			cat_container = cats.ItemContainerGenerator.ContainerFromItem(cat)
			child_0 = VisualTreeHelper.GetChild(cat_container,0)
			child_1 = VisualTreeHelper.GetChild(child_0,0)
			dataTemplate = child_1.ContentTemplate
			cat_exp = dataTemplate.FindName("cat_expander",child_1)
			if i in expanded_cats:
				cat_exp.IsExpanded = True
				list_pars = dataTemplate.FindName("list_pars",child_1)
				j = 0
				list_pars.UpdateLayout()
				for par in list_pars.ItemsSource:
					par_container = list_pars.ItemContainerGenerator.ContainerFromItem(par)
					child_0 = VisualTreeHelper.GetChild(par_container,0)
					child_1 = VisualTreeHelper.GetChild(child_0,0)
					dataTemplate = child_1.ContentTemplate
					par_exp = dataTemplate.FindName("par_expander",child_1)
					if i * 10 + j in expanded_pars:
						par_exp.IsExpanded = True
					j += 1
			i += 1
		self.myScroll.ScrollToVerticalOffset(pos_scroller)

	def selected_par(self, sender, args):
		self.fun_par(True)
	def unselected_par(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self.fun_par(False)
	def indeter_par(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self.fun_par(False)
	
	def and_Checked(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self.and_checked = True

	def or_Checked(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self.and_checked = False

	def check_cat_expanded(self):
		try:
			all_cats = self.list_cats.ItemsSource 
			cats = self.list_cats
			self.list_cats.UpdateLayout()
			for cat in cats.ItemsSource:
				cat_container = cats.ItemContainerGenerator.ContainerFromItem(cat)
				child_0 = VisualTreeHelper.GetChild(cat_container,0)
				child_1 = VisualTreeHelper.GetChild(child_0,0)
				dataTemplate = child_1.ContentTemplate
				cat_expander = dataTemplate.FindName("cat_expander",child_1)
				if cat_expander.IsExpanded:
					return True
			
			return False
		except:
			return False
	def check_par_expanded(self):
		try:
			all_cats = self.list_cats.ItemsSource 
			cats = self.list_cats
			for cat in cats.ItemsSource:
				cat_container = cats.ItemContainerGenerator.ContainerFromItem(cat)
				child_0 = VisualTreeHelper.GetChild(cat_container,0)
				child_1 = VisualTreeHelper.GetChild(child_0,0)
				dataTemplate = child_1.ContentTemplate
				cat_expander = dataTemplate.FindName("cat_expander",child_1)
				if cat_expander.IsExpanded:
					list_pars = dataTemplate.FindName("list_pars",child_1)
					ret_pars_elements = []
					ind_par = 0
					list_pars.UpdateLayout()
					for par in list_pars.ItemsSource:
						ret_par_elements = []
						par_container = list_pars.ItemContainerGenerator.ContainerFromItem(par)
						child_0 = VisualTreeHelper.GetChild(par_container,0)
						child_1 = VisualTreeHelper.GetChild(child_0,0)
						dataTemplate = child_1.ContentTemplate
						par_expander = dataTemplate.FindName("par_expander",child_1)
						if par_expander.IsExpanded:
							return True
			
			return False
		except:
			return False

	def fun_el(self,checked):
		"""Mark selected checkboxes as unchecked."""
		# save the expanded categories
		i = 0
		expanded_cats = []
		expanded_pars = []
		# set the new states
		all_cats = self.list_cats.ItemsSource 
		cats = self.list_cats
		self.list_cats.UpdateLayout()
		for cat in cats.ItemsSource:
			cat_container = cats.ItemContainerGenerator.ContainerFromItem(cat)
			child_0 = VisualTreeHelper.GetChild(cat_container,0)
			child_1 = VisualTreeHelper.GetChild(child_0,0)
			dataTemplate = child_1.ContentTemplate
			cat_exp = dataTemplate.FindName("cat_expander",child_1)
			if cat_exp.IsExpanded == True:
				list_pars = dataTemplate.FindName("list_pars",child_1)
				expanded_cats.append(i)
				j = 0
				checked_pars = 0
				par_None = 0
				list_pars.UpdateLayout()
				for par in list_pars.ItemsSource:
					par_container = list_pars.ItemContainerGenerator.ContainerFromItem(par)
					child_0 = VisualTreeHelper.GetChild(par_container,0)
					child_1 = VisualTreeHelper.GetChild(child_0,0)
					dataTemplate = child_1.ContentTemplate
					par_checkBox = dataTemplate.FindName("par_checkBox",child_1)
					par_exp = dataTemplate.FindName("par_expander",child_1)
					par.par_state = False
					if par_exp.IsExpanded == True:
						expanded_pars.append(i * 10 + j)
						checked_els = 0
						list_els = dataTemplate.FindName("list_els",child_1)
						if len(list_els.SelectedItems) > 0:
							for el in list_els.SelectedItems:
								el.el_state = checked
							for el in list_els.ItemsSource:
								if el.el_state == True:
									checked_els += 1
						else:
							list_els.UpdateLayout()
							for el in list_els.ItemsSource:
								par_container = list_els.ItemContainerGenerator.ContainerFromItem(el)
								child_0 = VisualTreeHelper.GetChild(par_container,0)
								child_1 = VisualTreeHelper.GetChild(child_0,0)
								dataTemplate = child_1.ContentTemplate
								el_checkBox = dataTemplate.FindName("el_checkBox",child_1)
								if el_checkBox.IsChecked == True:
									checked_els += 1
									el.el_state = True
								else:
									el.el_state = False
						if checked_els == len(list_els.ItemsSource):
							par.par_state = True
							checked_pars += 1
						elif checked_els > 0:
							par.par_state = None
							par_None += 1

					else:
						if par_checkBox.IsChecked == True:
							checked_pars += 1
							par.par_state = True
					j += 1
				cat.cat_state = False
				if checked_pars == len(list_pars.ItemsSource):
					cat.cat_state = True
				elif checked_pars > 0 or par_None > 0:
					cat.cat_state = None

			i += 1
		pos_scroller = self.myScroll.VerticalOffset
		self.list_cats.ItemsSource = []
		self.list_cats.ItemsSource = all_cats
		# reset the expanded categories
		all_cats = self.list_cats.ItemsSource 
		cats = self.list_cats
		i = 0
		cats.UpdateLayout()
		for cat in cats.ItemsSource:
			cat_container = cats.ItemContainerGenerator.ContainerFromItem(cat)
			child_0 = VisualTreeHelper.GetChild(cat_container,0)
			child_1 = VisualTreeHelper.GetChild(child_0,0)
			dataTemplate = child_1.ContentTemplate
			cat_exp = dataTemplate.FindName("cat_expander",child_1)
			if i in expanded_cats:
				cat_exp.IsExpanded = True
				list_pars = dataTemplate.FindName("list_pars",child_1)
				j = 0
				list_pars.UpdateLayout()
				for par in list_pars.ItemsSource:
					par_container = list_pars.ItemContainerGenerator.ContainerFromItem(par)
					child_0 = VisualTreeHelper.GetChild(par_container,0)
					child_1 = VisualTreeHelper.GetChild(child_0,0)
					dataTemplate = child_1.ContentTemplate
					par_exp = dataTemplate.FindName("par_expander",child_1)
					if i * 10 + j in expanded_pars:
						par_exp.IsExpanded = True
					j += 1
			i += 1
		self.myScroll.ScrollToVerticalOffset(pos_scroller)

	def selected_el(self, sender, args):
		self.fun_el(True)

	def unselected_el(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self.fun_el(False)

	def select_els(self,sender,args):
		"""Mark selected checkboxes as unchecked."""
		# selected elements
		returned_els = []
		all_cats = self.list_cats.ItemsSource 
		cats = self.list_cats
		cats.UpdateLayout()

		for cat in cats.ItemsSource:
			cat_container = cats.ItemContainerGenerator.ContainerFromItem(cat)
			child_0 = VisualTreeHelper.GetChild(cat_container,0)
			child_1 = VisualTreeHelper.GetChild(child_0,0)
			dataTemplate = child_1.ContentTemplate
			cat_checkBox = dataTemplate.FindName("cat_checkBox",child_1)
			if cat_checkBox.IsChecked == True:
				for el in cat.cat_els:
					returned_els.Add(el.Id)
				continue
			if cat_checkBox.IsChecked == None:
				list_pars = dataTemplate.FindName("list_pars",child_1)
				ret_pars_elements = []
				ind_par = 0
				list_pars.UpdateLayout()
				for par in list_pars.ItemsSource:
					ret_par_elements = []
					par_container = list_pars.ItemContainerGenerator.ContainerFromItem(par)
					child_0 = VisualTreeHelper.GetChild(par_container,0)
					child_1 = VisualTreeHelper.GetChild(child_0,0)
					dataTemplate = child_1.ContentTemplate
					par_checkBox = dataTemplate.FindName("par_checkBox",child_1)
					if par_checkBox.IsChecked == True:
						for el in par.ret_elements:
							ret_par_elements.append(el.Id)
					if par_checkBox.IsChecked == None:
						list_els = dataTemplate.FindName("list_els",child_1)
						list_els.UpdateLayout()
						for el in list_els.ItemsSource:
							par_container = list_els.ItemContainerGenerator.ContainerFromItem(el)
							child_0 = VisualTreeHelper.GetChild(par_container,0)
							child_1 = VisualTreeHelper.GetChild(child_0,0)
							dataTemplate = child_1.ContentTemplate
							el_checkBox = dataTemplate.FindName("el_checkBox",child_1)
							if el_checkBox.IsChecked == True:
								for e in par.ret_elements:
									if e.par_value:
										if e.par_value == el.par_value:
											ret_par_elements.append(e.Id)
					

					if len(ret_par_elements) == 0:
						continue

					if self.and_checked:
						
						if len(ret_pars_elements) == 0:
							for el in ret_par_elements:
								ret_pars_elements.Add(el)
						else:
							new_els = []
							for el in ret_par_elements:
								if el in ret_pars_elements:
									new_els.Add(el)
							ret_pars_elements = []
							ret_pars_elements = new_els
					else:
						for el in ret_par_elements:
							if not el in ret_pars_elements:
								ret_pars_elements.Add(el)
					
				for el in ret_pars_elements:
					if not el in returned_els:
						returned_els.Add(el)

		if len(returned_els) > 0:
			__revit__.ActiveUIDocument.Selection.SetElementIds(List[DB.ElementId](returned_els))
			self.Close()
		else:
			MessageBox.Show("Выделенных элементов нет !!","ОШИБКА!")

			
		# close print window
		output = pyrevit.output.get_output()
		output.close()

	def select_allcats(self,sender,args):

		self.Close()
		loadCategories(False)

	def cancel(self,sender,args):
		output = pyrevit.output.get_output()
		output.close()
		self.Close()

	def MyMouseWheelV(self,sender,args):
		self.myScroll.ScrollToVerticalOffset(self.myScroll.VerticalOffset - args.Delta)
		
def GroupByParameter(lst, func):
	res = {}
	for el in lst:
		key = func(el)
		if key in res:
			res[key].append(el)
		else:
			res[key] = [el]
	return res

class category1(object):
	def __init__(self,cat):

		self.cat_Name = cat
		self.cat_state = False
		self.IsGroovy = False

	def get_elements(self,els):
		cat_els = []
		paramters_get = True
		for el in els:
			if el.Category.Name == self.cat_Name:
				cat_els.append(el)
		return  cat_els

class category(object):
	def __init__(self,cat,els):
		self.elments = els
		self.cat_Name = cat
		x = self.get_paramters(els)
		pars_0 = x[0]
		self.cat_els = x[1]
		pars_1 = [paramter(par,self.cat_els) for par in pars_0]
		pars_2 = [par for par in pars_1 if  not len(par.elements) == 0]
		self.paramters = sorted(pars_2, key= lambda x:x.par_name)
		self.cat_state = False
		self.IsGroovy = False

	def get_paramters(self,els):
		pars = []
		cat_els = []
		paramters_get = True
		for el in els:
			if el.Category.Name == self.cat_Name:
				if paramters_get:
					pars = el.Parameters
					paramters_get = False
				cat_els.append(el)
		cat_pars = pars
		return cat_pars, cat_els




class paramter(object):
	def __init__(self, par, els):
		self.paramter = par
		self.par_name = par.Definition.Name
		all_els = []
		self.inits_els = []
		par_els = []
		els_par_value = []
		par_u_els = []
		for el in els:
			el_par = element(el, par)
			if not el_par.par_value == None:
				par_els.append(el_par)
				if not el_par.par_value in els_par_value :
					els_par_value.append(el_par.par_value)
					par_u_els.append(el_par)

		self.elements = par_u_els
		self.ret_elements = par_els

		self.par_state = False


class element(object):
	def __init__(self, obj, par):

		self.__obj = obj
		self.Id = obj.Id
		self.par_value = ''
		if obj.LookupParameter(par.Definition.Name):
		 	self.par_value = obj.LookupParameter(par.Definition.Name).AsString()
			if self.par_value == None:
				self.par_value = obj.LookupParameter(par.Definition.Name).AsValueString()
		self.el_state = False

class UseSelectedItems(forms.TemplateUserInputWindow):
	xaml_source = op.join(op.dirname(__file__),'UseSelectedItems.xaml')
	
	def _setup(self, **kwargs):
		self.Title = "Выделенные элементы"
	

	def check_select(self, sender, args):
		"""Handle select button click."""
		self.response = {'checked':True}
		self.Close()
	
	def uncheck_select(self, sender, args):
		"""Handle select button click."""
		self.response = {'checked':False}
		self.Close()

def loadCategories(first_load):
	doc = __revit__.ActiveUIDocument.Document

	uidoc = __revit__.ActiveUIDocument
	v = __revit__.ActiveUIDocument.ActiveView 

	view = __revit__.ActiveUIDocument.ActiveGraphicalView 
	cats_2_remove = ['<Эскиз лестницы/пандуса: Граница>','<Эскиз>','Базовая точка проекта',
					'Варианты нагружений на конструкцию','Общая площадка',
					'Классификации нагрузок в электросетях','Наборы характеристик материала',
					'Определения коэффициентов спроса электрических нагрузок','Основные горизонтали',
					'Параметры назначения пространства',
					'Параметры типа здания','Сведения о проекте','Сетка рабочей плоскости',
					'Схема цветового обозначения''Схемы зонирования','Точка съемки','Траектория солнца',
					'Шаблоны принципиальной схемы щита/панели - Распределительный щит',
					'Шаблоны принципиальной схемы щита/панели - Силовой щит',
					'Шаблоны принципиальной схемы щита/панели - Телекоммуникационный щит',
					'Элемент параметра классификации нагрузок в электросетях']
	sel_elements = [ el for el in __revit__.ActiveUIDocument.Selection.GetElementIds()]
	if len(sel_elements) > 0:
		if first_load:
			ops = []
			res = UseSelectedItems.show(ops, button_name='Рассчитать')
			if res:
				SELECTION = res['checked']
			else:
				raise SystemExit(1)
			if not SELECTION:
				els = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
			else:
				elements = [ el for el in __revit__.ActiveUIDocument.Selection.GetElementIds()]
				els = []
				for id in elements:
					el = uidoc.Document.GetElement(id)
					els.append(el)
		else:
			elements = [ el for el in __revit__.ActiveUIDocument.Selection.GetElementIds()]
			els = []
			for id in elements:
				el = uidoc.Document.GetElement(id)
				els.append(el)

	else:
		els = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()

	# get the Categories
	categories = []
	els_hasCat = []

	for el in els :
		cat = el.Category
		if not cat is None:
			els_hasCat.append(el)
			cat_Name = el.Category.Name
			if (not cat_Name in categories) and (not cat_Name in cats_2_remove):
				categories.append(cat_Name)


	list_cats = [category1(cat) for cat in categories]
	sorted_cats = sorted(list_cats, key= lambda x:x.cat_Name)
	QuickSelectWindow('quickSelect_main.xaml',list = sorted_cats, Elements = els_hasCat).ShowDialog()

def main():
	loadCategories(True)

if __name__ == "__main__":
	main()