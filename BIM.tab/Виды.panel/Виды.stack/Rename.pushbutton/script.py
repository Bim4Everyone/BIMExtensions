# -*- coding: utf-8 -*-

import os.path as op
from operator import itemgetter
import clr, sys, os

clr.AddReference("System.Windows.Forms")
clr.AddReference("System.Drawing")
from System.Windows.Forms import MessageBox
from System.Drawing import Point
#импорт из библиотеки Autodesk
from Autodesk.Revit.DB import ViewSheet, FilteredElementCollector, BuiltInCategory, BuiltInParameter, Transaction, TransactionGroup
#ипорт из библиотеки pyRevit
from pyrevit import forms
from pyrevit.framework import Controls
from pyrevit.forms import DEFAULT_INPUTWINDOW_WIDTH, DEFAULT_INPUTWINDOW_HEIGHT, SelectFromList, TemplateUserInputWindow
from pyrevit.coreutils import Timer

#Speech библиотека
class TemplateWindow(TemplateUserInputWindow):

	def __init__(self, context, title, **kwargs):
		forms.WPFWindow.__init__(self, op.join(op.dirname(__file__), self.xaml_source))
		self.Title = title

		self._context = context
		self.response = None
		self.PreviewKeyDown += self.handle_input_key

		self._setup(**kwargs)
	
	@classmethod
	def show(cls, context,
			title='User Input',
			width=DEFAULT_INPUTWINDOW_WIDTH,
			height=DEFAULT_INPUTWINDOW_HEIGHT, **kwargs):
		"""Show user input window.

		Args:
			context (any): window context element(s)
			title (type): window title
			width (type): window width
			height (type): window height
			**kwargs (type): other arguments to be passed to window
		"""
		dlg = cls(context, title, **kwargs)
		dlg.ShowDialog()
		return dlg.response

		
class Renamer(TemplateWindow):

	xaml_source = os.path.dirname(os.path.abspath(__file__))+'\\RenameThree.xaml'
	
	def _setup(self, **kwargs):
		self.hide_element(self.clrsuffix_b)
		self.hide_element(self.clrprefix_b)
		self.hide_element(self.clrold_b)
		self.hide_element(self.clrnew_b)
		self.clear_suffix(None, None)
		self.clear_prefix(None, None)
		self.clear_old(None, None)
		self.clear_new(None, None)
		self.prefix.Focus()
		
		self.old_name.IsEnabled = False
		self.new_name.IsEnabled = False
		
		self.action = True
		
		button_name = kwargs.get('button_name', None)
		if button_name:
			self.select_b.Content = button_name
		


	def button_select(self, sender, args):
		if self.suffix.Text:
			suffix = self.suffix.Text
		else:
			suffix = ""
		if self.prefix.Text:
			prefix = self.prefix.Text
		else:
			prefix = ""
		if self.old_name.Text:
			old_name = self.old_name.Text
		else:
			old_name = ""
		if self.new_name.Text:
			new_name = self.new_name.Text
		else:
			new_name = ""
		self.response = {"action": self.action,
						"suffix": suffix,
						"prefix": prefix,
						"prefix_checked": self.prefix_rep.IsChecked,
						"suffix_checked": self.suffix_rep.IsChecked,
						"new_name": new_name,
						"old_name": old_name}
		self.Close()
	
	def rename_Checked(self, sender, args):
		self.action = False
	
		self.old_name.IsEnabled = True
		self.new_name.IsEnabled = True
		self.clrold_b.IsEnabled = True
		self.clrnew_b.IsEnabled = True
		
		self.prefix.IsEnabled = False
		self.suffix.IsEnabled = False
		self.clrsuffix_b.IsEnabled = False
		self.clrprefix_b.IsEnabled = False
		self.suffix_rep.IsEnabled = False
		self.prefix_rep.IsEnabled = False
		
	def presuf_Checked(self, sender, args):
		self.action = True
		
		self.old_name.IsEnabled = False
		self.new_name.IsEnabled = False
		self.clrold_b.IsEnabled = False
		self.clrnew_b.IsEnabled = False
		
		self.prefix.IsEnabled = True
		self.suffix.IsEnabled = True
		self.clrsuffix_b.IsEnabled = True
		self.clrprefix_b.IsEnabled = True
		self.suffix_rep.IsEnabled = True
		self.prefix_rep.IsEnabled = True
	
	def suffix_txt_changed(self, sender, args):
		"""Handle text change in search box."""
		if self.suffix.Text == '':
			self.hide_element(self.clrsuffix_b)
		else:
			self.show_element(self.clrsuffix_b)
	
	def prefix_txt_changed(self, sender, args):
		"""Handle text change in search box."""
		if self.prefix.Text == '':
			self.hide_element(self.clrprefix_b)
		else:
			self.show_element(self.clrprefix_b)

	def old_txt_changed(self, sender, args):
		"""Handle text change in search box."""
		if self.old_name.Text == '':
			self.hide_element(self.clrold_b)
		else:
			self.show_element(self.clrold_b)
	
	def new_txt_changed(self, sender, args):
		"""Handle text change in search box."""
		if self.new_name.Text == '':
			self.hide_element(self.clrnew_b)
		else:
			self.show_element(self.clrnew_b)
		
	def clear_suffix(self, sender, args):
		"""Clear search box."""
		self.suffix.Text = ''
		self.suffix.Clear()
		self.suffix.Focus
		
	def clear_prefix(self, sender, args):
		"""Clear search box."""
		self.prefix.Text = ''
		self.prefix.Clear()
		self.prefix.Focus

	def clear_old(self, sender, args):
		"""Clear search box."""
		self.old_name.Text = ''
		self.old_name.Clear()
		self.old_name.Focus
		
	def clear_new(self, sender, args):
		"""Clear search box."""
		self.new_name.Text = ''
		self.new_name.Clear()
		self.new_name.Focus
		

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
selections = [ doc.GetElement( elId ) for elId in __revit__.ActiveUIDocument.Selection.GetElementIds() ]
res = Renamer.show([], title='Переименование', button_name='Ок')

tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()

for selection in selections:
	
	if res:
		oldName = selection.Name
		newName = ''
		try:
			newName = oldName[:oldName.index('копия')-1]
		except:
			newName = oldName
		if res['action']:
			if res['suffix_checked']:
				temp_name = newName[::-1]
				temp_name = temp_name[temp_name.index(' ')+1:]
				newName = temp_name[::-1]
			if res['suffix']:
				newName += ' ' + res['suffix']
			if res['prefix_checked']:
				newName = newName[newName.index(' ')+1:]
			if res['prefix']:
				newName = res['prefix'] + ' ' + newName
		else:
			if res['old_name']:
				drop = newName.partition(res['old_name'])
				newName = drop[0] + res['new_name'] + drop[2]
			
		selection.Name = newName
		
t.Commit()
tg.Assimilate()