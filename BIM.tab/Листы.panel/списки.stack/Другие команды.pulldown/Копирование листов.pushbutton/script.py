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

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

import re
from System.Windows.Data import CollectionViewSource, PropertyGroupDescription
from System.Windows.Media import VisualTreeHelper

import dosymep.Revit
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
view = __revit__.ActiveUIDocument.ActiveGraphicalView
view = doc.ActiveView


def sort_fun(str):
    num = [int(x) for x in re.findall(r'\d+', str)]
    if len(num) > 0:
        return num[0]
    else:
        return 1000


def GroupByParameter(lst, func):
    res = {}
    for el in lst:
        key = func(el)
        if key in res:
            res[key].append(el)
        else:
            res[key] = [el]
    return res


def list_albums(lists):
    albums = []
    for li in lists:
        if not li in albums:
            albums.append(li)
    return albums


class SheetOption(object):
    def __init__(self, obj, state=False):
        self.state = state
        self.name = "{} - {}".format(obj.SheetNumber, obj.Name)
        self.Number = obj.SheetNumber
        self.speech_album = obj.GetParamValueOrDefault(SharedParamsConfig.Instance.AlbumBlueprints)

        self.obj = obj

        def __nonzero__(self):
            return self.state

        def __str__(self):
            return self.name


class SelectLists(forms.WPFWindow):
    def __init__(self, xaml_file_name, **kwargs):
        forms.WPFWindow.__init__(self, xaml_file_name)

        self.checked_only = kwargs.get('checked_only', True)
        button_name = kwargs.get('button_name', None)
        self._context = kwargs.get('list', None)
        self.views = kwargs.get('views', None)
        if button_name:
            self.select_b.Content = button_name
        self.hide_element(self.clrsuffix_b)
        self.hide_element(self.clrprefix_b)
        self.list_lb.SelectionMode = Controls.SelectionMode.Extended

        self._tbAlbumBlueprints.Text = SharedParamsConfig.Instance.AlbumBlueprints.Name

        for i in kwargs['views']:
            self.purpose.AddText(i)
        for i in kwargs['speech_albums']:
            self.sp_albums.AddText(i)
        # self._verify_context()
        self._list_options()

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
        return self.list_lb.ItemsSource

    @sheet_list.setter
    def sheet_list(self, value):
        self.list_lb.ItemsSource = value

    @property
    def selected_sheets(self):
        return self.list_lb.SelectedItems

    def _verify_context(self):
        new_context = []
        for item in self.sheet_list:
            if not hasattr(item, 'state'):
                new_context.append(BaseCheckBoxItem(item))
            else:
                new_context.append(item)

        self.sheet_list = new_context

    def _list_options(self, checkbox_filter=None):
        if checkbox_filter:
            checkbox_filter = checkbox_filter.lower()
            self.list_lb.ItemsSource = \
                [checkbox for checkbox in self._context
                 if checkbox_filter in checkbox.name.lower()]
        else:
            self.list_lb.ItemsSource = self._context

    def dtGrid_ScrollChanged(self, sender, args):
        self.xx = False

    def _set_states(self, state=True, selected=False):
        if self.selected_sheets:
            if self.sheet_list:
                all_items = self.sheet_list
                current_list = self.selected_sheets
                for it in current_list:
                    it.state = state

                self.sheet_list = []
                self.sheet_list = all_items

    def check_selected(self, sender, args):
        """Mark selected checkboxes as checked."""
        border = VisualTreeHelper.GetChild(self.list_lb, 0)
        self.scrollViewer = border.Child
        offset = self.scrollViewer.VerticalOffset
        self._set_states(state=True, selected=True)
        self.scrollViewer.ScrollToVerticalOffset(offset)

    def uncheck_selected(self, sender, args):
        """Mark selected checkboxes as unchecked."""
        border = VisualTreeHelper.GetChild(self.list_lb, 0)
        self.scrollViewer = border.Child
        offset = self.scrollViewer.VerticalOffset
        self._set_states(state=False, selected=True)
        self.scrollViewer.ScrollToVerticalOffset(offset)

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

    def button_select(self, sender, args):
        """Handle select button click."""
        if self.checked_only:
            sheets = [x for x in self._context if x.state]
        else:
            sheets = self._context

        if self.copyViews.IsChecked:
            copyViews = True
        else:
            copyViews = False
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

        if self.sp_albums.Text:
            speech_album = self.sp_albums.Text
        else:
            speech_album = ""
        self.response = {"copyViews": copyViews,
                         "suffix": suffix,
                         "prefix": prefix,
                         "purpose": purpose,
                         "speech_album": speech_album,
                         "sheets": sheets}
        self.Close()


class ViewSheetDuplicater:
    def __init__(self):
        self.doc = __revit__.ActiveUIDocument.Document
        self.sheets = FilteredElementCollector(doc) \
            .OfClass(ViewSheet) \
            .ToElements()

        views = self.GetViewPurposeList()
        speech_albums = self.GetSpeechAlbumList()
        window = SelectLists('SelectLists.xaml',
                             title="Копировние Листов",
                             button_name="Копировать",
                             list=self.list_shts(),
                             views=views,
                             speech_albums=speech_albums)
        window.ShowDialog()

        if hasattr(window, 'response'):
            res = window.response
            self.purpose = res["purpose"]
            self.prefix = res["prefix"]
            self.suffix = res["suffix"]
            self.copyViews = res["copyViews"]
            self.speech_album = res["speech_album"]
            for item in res["sheets"]:
                sheet = item.obj
                sheetId = sheet.Id
                titleType = self._GetTitleType(sheetId)
                viewports = self._GetViewports(sheetId)
                details = self._GetDetails(sheetId)
                textNotes = self._GetTextNotes(sheetId)
                scheduleSheets = self._GetScheduleSheets(sheetId)
                self.Duplicate(sheet, titleType, viewports, details, textNotes, scheduleSheets)

    def list_shts(self):
        sheets = FilteredElementCollector(doc) \
            .OfClass(ViewSheet) \
            .ToElements()

        list_sheets = [SheetOption(sh) for sh in sheets]
        sorted_sheets = []
        albums = [x.speech_album for x in list_sheets]
        speech_albums = list_albums(albums)
        grouped_sheets = GroupByParameter(list_sheets, func=lambda x: x.sheet_album)
        for album in speech_albums:
            ll = sorted(grouped_sheets[album], key=lambda x: sort_fun(x.Number))
            for z in ll:
                sorted_sheets.append(z)

        view = CollectionViewSource.GetDefaultView(sorted_sheets)
        groupDescription = PropertyGroupDescription('speech_album')
        view.GroupDescriptions.Add(groupDescription)
        return view

    def Duplicate(self, sheet, titleType, viewports, details, textNotes, scheduleSheets):
        tg = TransactionGroup(doc, "Duplicateing Sheet")
        tg.Start()
        t = Transaction(doc, "Duplicateing Sheet")
        t.Start()

        duplicatedSheet = ViewSheet.Create(doc, titleType if titleType else ElementId.InvalidElementId)
        duplicatedSheet.Name = sheet.Name
        duplicatedSheet.SetParamValue(SharedParamsConfig.Instance.AlbumBlueprints, self.speech_album)
        if self.copyViews:
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
                    viewName = originalView.Name
                    duplicatedViewId = originalView.Duplicate(ViewDuplicateOption.WithDetailing)
                    duplicatedView = doc.GetElement(duplicatedViewId)

                    duplicatedViewName = []
                    if self.prefix: duplicatedViewName.append(self.prefix)
                    duplicatedViewName.append(viewName)
                    if self.suffix: duplicatedViewName.append(self.suffix)
                    duplicatedViewName = " ".join(duplicatedViewName)

                    if len(duplicatedViewName) > len(viewName):
                        duplicatedView.Name
                    try:
                        duplicatedView.SetParamValue(ProjectParamsConfig.Instance.ViewGroup, self.purpose)
                    except:
                        duplicatedView.ViewTemplateId = ElementId(-1)
                        duplicatedView.SetParamValue(ProjectParamsConfig.Instance.ViewGroup, self.purpose)

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
            param = view.GetParamValueOrDefault(ProjectParamsConfig.Instance.ViewGroup)
            if param not in purpose:
                purpose.append(param)
        return purpose

    def GetSpeechAlbumList(self):
        sheets = FilteredElementCollector(doc) \
            .OfClass(ViewSheet) \
            .ToElements()
        speechAlbumList = []
        for sh in sheets:
            if sh.GetParamValueOrDefault(SharedParamsConfig.Instance.AlbumBlueprints):
                param = sh.GetParamValueOrDefault(SharedParamsConfig.Instance.AlbumBlueprints)
                if param not in speechAlbumList:
                    speechAlbumList.append(param)
        return speechAlbumList

    def _GetTitleType(self, sheetId):
        collector = FilteredElementCollector(self.doc)
        collector.OfClass(FamilyInstance)
        collector.OfCategory(BuiltInCategory.OST_TitleBlocks)
        collector.OwnedByView(sheetId)
        titleType = None
        for element in collector:
            if element.OwnerViewId == sheetId:
                titleType = element.GetTypeId()
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

        details = List[ElementId]([x.Id for x in collector if isinstance(x, DetailLine) \
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
        collector.OwnedByView(sheetId)
        scheduleSheets = List[ElementId]([x.Id for x in collector])

        return scheduleSheets


duplicater = ViewSheetDuplicater()