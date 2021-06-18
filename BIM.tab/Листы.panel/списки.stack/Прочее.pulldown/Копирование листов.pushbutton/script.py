# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
import clr
import math
import collections
from shutil import copyfile
from pyrevit import forms
from pyrevit.framework import Controls
clr.AddReference('System')
clr.AddReference('System.IO')
clr.AddReference("System.Windows.Forms")
from System.IO import FileInfo
from System.Windows.Forms import MessageBox, SaveFileDialog, DialogResult
from System.Collections.Generic import List
from Autodesk.Revit.DB import DetailNurbSpline, CurveElement, ElementTransformUtils, \
                              DetailLine, View, ViewDuplicateOption, XYZ, LocationPoint, \
                              TransactionGroup, Transaction, FilteredElementCollector, \
                              ElementId, BuiltInCategory, FamilyInstance, ViewDuplicateOption, \
                              ViewSheet, FamilySymbol, Viewport, DetailEllipse, DetailArc, TextNote, \
                              ScheduleSheetInstance
from Autodesk.Revit.Creation import ItemFactoryBase
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.UI import RevitCommandId, PostableCommand

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
view = __revit__.ActiveUIDocument.ActiveGraphicalView
view = doc.ActiveView

def alert(msg):
    MessageBox.Show(msg)

class SheetOption(object):
    def __init__(self, obj, state=False):
        self.state = state
        self.name = "{} - {}".format(obj.SheetNumber, obj.Name)
        self.obj = obj
        def __nonzero__(self):
            return self.state
        def __str__(self):
            return self.name

class SelectLevelFrom(forms.TemplateUserInputWindow):
    xaml_source = op.join(op.dirname(__file__),'SelectFromCheckboxes.xaml')
    
    def _setup(self, **kwargs):
        self.checked_only = kwargs.get('checked_only', True)
        button_name = kwargs.get('button_name', None)
        self._context = kwargs.get('list', None)
        self.views = kwargs.get('views', None)
        if button_name:
            self.select_b.Content = button_name
        self.hide_element(self.clrsuffix_b)
        self.hide_element(self.clrprefix_b)
        self.list_lb.SelectionMode = Controls.SelectionMode.Extended

        for i in kwargs['views']:
            self.purpose.AddText(i)
        
        #for i in range(1,4):
        #	self.purpose.AddText(str(i))

        self._verify_context()
        self._list_options()

    def button_select(self, sender, args):
        """Handle select button click."""
        if self.checked_only:
            sheets = [x for x in self._context if x.state]
        else:
            sheets = self._context

        if self.suffix.Text:
            suffix = self.suffix.Text
        else:
            suffix = ""
        if self.prefix.Text:
            prefix = self.prefix.Text
        else:
            prefix = ""
        if self.purpose.Text:
            purpose = self.purpose.Text
        else:
            purpose = ""
        self.response = {"suffix": suffix,
                        "prefix": prefix,
                        "purpose": purpose,
                        "sheets": sheets}
        self.Close()

    def _verify_context(self):
        new_context = []
        for item in self._context:
            if not hasattr(item, 'state'):
                new_context.append(BaseCheckBoxItem(item))
            else:
                new_context.append(item)

        self._context = new_context

    def _list_options(self, checkbox_filter=None):
        if checkbox_filter:
            checkbox_filter = checkbox_filter.lower()
            self.list_lb.ItemsSource = \
                [checkbox for checkbox in self._context
                if checkbox_filter in checkbox.name.lower()]
        else:
            self.list_lb.ItemsSource = self._context

    def _set_states(self, state=True, flip=False, selected=False):
        all_items = self.list_lb.ItemsSource
        if selected:
            current_list = self.list_lb.SelectedItems
        else:
            current_list = self.list_lb.ItemsSource
        for checkbox in current_list:
            if flip:
                checkbox.state = not checkbox.state
            else:
                checkbox.state = state

        # push list view to redraw
        self.list_lb.ItemsSource = None
        self.list_lb.ItemsSource = all_items

    def check_selected(self, sender, args):
        """Mark selected checkboxes as checked."""
        self._set_states(state=True, selected=True)

    def uncheck_selected(self, sender, args):
        """Mark selected checkboxes as unchecked."""
        self._set_states(state=False, selected=True)

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



class ViewSheetDuplicater:
    def __init__(self):
        self.doc = __revit__.ActiveUIDocument.Document
        self.sheets = FilteredElementCollector(doc) \
                        .OfClass(ViewSheet) \
                        .ToElements()

        views = self.GetViewPurposeList()
        
        list = [SheetOption(x) for x in self.sheets]
        list = sorted(list, key=lambda x: x.name)

        response = SelectLevelFrom.show([], 
                                        title= "Копировние Листов", 
                                        button_name="Ок", 
                                        list=list, 
                                        views=views)
        if response:
            self.purpose = response["purpose"]
            self.prefix = response["prefix"]
            self.suffix = response["suffix"]
            
            for item in response["sheets"]:
                sheet = item.obj
                sheetId = sheet.Id

                titleType = self._GetTitleType(sheetId)
                viewports = self._GetViewports(sheetId)
                details = self._GetDetails(sheetId)
                textNotes = self._GetTextNotes(sheetId)
                scheduleSheets = self._GetScheduleSheets(sheetId)
                self.Duplicate(sheet, titleType, viewports, details, textNotes, scheduleSheets)


    def Duplicate(self, sheet, titleType, viewports, details, textNotes, scheduleSheets):
        tg = TransactionGroup(doc, "Duplicateing Sheet")
        tg.Start()
        t = Transaction(doc, "Duplicateing Sheet")
        t.Start()

        duplicatedSheet = ViewSheet.Create(doc, titleType)
        duplicatedSheet.Name = sheet.Name
        for viewport in viewports:
            viewportTypeId = viewport.GetTypeId()
            if Viewport.CanAddViewToSheet(self.doc, 
                                          duplicatedSheet.Id, 
                                          viewport.ViewId):
                duplicatedViewport = Viewport.Create(self.doc, 
                                                    duplicatedSheet.Id, 
                                                    viewport.ViewId, 
                                                    viewport.GetBoxCenter())
                duplicatedViewport.ChangeTypeId(viewportTypeId)
                
                                
            else:
                originalView = doc.GetElement(viewport.ViewId)
                viewName = originalView.LookupParameter("Имя вида").AsString()
                duplicatedViewId = originalView.Duplicate(ViewDuplicateOption.WithDetailing)
                duplicatedView = doc.GetElement(duplicatedViewId)
                
                duplicatedViewName = []
                if self.prefix: duplicatedViewName.append(self.prefix)
                duplicatedViewName.append(viewName)
                if self.suffix: duplicatedViewName.append(self.suffix)
                duplicatedViewName = " ".join(duplicatedViewName)
                
                if len(duplicatedViewName) > len(viewName):
                    duplicatedView.LookupParameter("Имя вида").Set(duplicatedViewName)
                try:
                    duplicatedView.LookupParameter("Назначение вида").Set(self.purpose)
                except:
                    parameter = duplicatedView.LookupParameter("Шаблон вида")
                    a = ElementId(-1)
                    parameter.Set(a)
                    duplicatedView.LookupParameter("Назначение вида").Set(self.purpose)

                duplicatedViewport = Viewport.Create(self.doc, 
                                                    duplicatedSheet.Id, 
                                                    duplicatedViewId, 
                                                    viewport.GetBoxCenter())
                duplicatedViewport.ChangeTypeId(viewportTypeId)
            

        if details:
            ElementTransformUtils.CopyElements(sheet, details, duplicatedSheet, None, None)
        if textNotes:
            ElementTransformUtils.CopyElements(sheet, textNotes, duplicatedSheet, None, None)
        if scheduleSheets:
            ElementTransformUtils.CopyElements(sheet, scheduleSheets, duplicatedSheet, None, None)


        t.Commit()
        tg.Assimilate()

        return duplicatedSheet


    def GetViewPurposeList(self):
        views = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Views).ToElements()
        purpose = []
        for view in views:
            parameter = view.LookupParameter("Назначение вида")
            if parameter is None: 
                alert("Параметр Назначение вида отсутствует")
                sys.exit().AsString()
            param = parameter
            if param not in purpose:
                purpose.append(param)
        return purpose
        


    def _GetTitleType(self, sheetId):
        collector = FilteredElementCollector(self.doc)
        collector.OfClass(FamilyInstance)
        collector.OfCategory(BuiltInCategory.OST_TitleBlocks)
        collector.OwnedByView(sheetId)
        titleType = None
        for element in collector:
            if element.OwnerViewId == sheetId:
                titleType = element.GetTypeId()
                #print element.GetTypeId()
                break

        return titleType


    def _GetViewports(self, sheetId):
        collector = FilteredElementCollector(doc)
        collector.OfClass(Viewport)
        collector.OwnedByView(sheetId)

        viewports = [x for x in collector]

        return viewports


    def _GetDetails(self, sheetId):
        collector = FilteredElementCollector(doc)
        collector.OfClass(CurveElement)
        collector.OwnedByView(sheetId)

        details = List[ElementId]([x.Id for x in collector if isinstance(x, DetailLine)\
                                                                or isinstance(x, DetailNurbSpline) \
                                                                or isinstance(x, DetailEllipse) \
                                                                or isinstance(x, DetailArc)])

        return details


    def _GetTextNotes(self, sheetId):
        collector = FilteredElementCollector(doc)
        collector.OfClass(TextNote)
        collector.OwnedByView(sheetId)
        
        textNotes = List[ElementId]([x.Id for x in collector])

        return textNotes

    def _GetScheduleSheets(self, sheetId):
        collector = FilteredElementCollector(doc)
        collector.OfClass(ScheduleSheetInstance)
        #print collector.GetElementCount()
        collector.OwnedByView(sheetId)
        #print collector.GetElementCount()
        scheduleSheets = List[ElementId]([x.Id for x in collector])

        return scheduleSheets

duplicater = ViewSheetDuplicater()