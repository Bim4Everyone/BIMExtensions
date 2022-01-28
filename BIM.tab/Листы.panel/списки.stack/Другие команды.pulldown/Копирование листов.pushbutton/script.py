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

from pyrevit import revit

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
view = __revit__.ActiveUIDocument.ActiveGraphicalView
view = doc.ActiveView

project_params = ProjectParameters.Create(app)
project_params.SetupRevitParams(doc,
                                ProjectParamsConfig.Instance.ViewGroup,
                                SharedParamsConfig.Instance.AlbumBlueprints)


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
        if li not in albums:
            albums.append(li)

    return albums


class SheetOption(object):
    def __init__(self, obj, state=False):
        self.state = state
        self.name = "{} - {}".format(obj.SheetNumber, obj.Name)
        self.Number = obj.SheetNumber
        self.sheet_album = obj.GetParamValueOrDefault(SharedParamsConfig.Instance.AlbumBlueprints)

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
        for i in kwargs['sheet_albums']:
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
            sheet_album = self.sp_albums.Text
        else:
            sheet_album = ""
        self.response = {"copyViews": copyViews,
                         "suffix": suffix,
                         "prefix": prefix,
                         "purpose": purpose,
                         "sheet_album": sheet_album,
                         "sheets": sheets}
        self.Close()


class ViewSheetDuplicater:
    def __init__(self):
        self.doc = __revit__.ActiveUIDocument.Document
        self.sheets = FilteredElementCollector(doc) \
            .OfClass(ViewSheet) \
            .ToElements()

        views = self.GetViewPurposeList()
        sheet_albums = self.GetSheetAlbumList()
        window = SelectLists('SelectLists.xaml',
                             title="Копирование Листов",
                             button_name="Копировать",
                             list=self.list_shts(),
                             views=views,
                             sheet_albums=sheet_albums)
        window.ShowDialog()

        if hasattr(window, 'response'):
            res = window.response
            self.purpose = res["purpose"]
            self.prefix = res["prefix"]
            self.suffix = res["suffix"]
            self.copyViews = res["copyViews"]
            self.sheet_album = res["sheet_album"]
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
        albums = [x.sheet_album for x in list_sheets]
        sheet_albums = list_albums(albums)
        grouped_sheets = GroupByParameter(list_sheets, func=lambda x: x.sheet_album)

        for album in sheet_albums:
            ll = sorted(grouped_sheets[album], key=lambda x: sort_fun(x.Number))
            for z in ll:
                sorted_sheets.append(z)

        view = CollectionViewSource.GetDefaultView(sorted_sheets)
        group_description = PropertyGroupDescription('sheet_album')
        view.GroupDescriptions.Add(group_description)

        return view

    def Duplicate(self, sheet, titleType, viewports, details, textNotes, scheduleSheets):
        with revit.Transaction("BIM: Копирование листов"):
            duplicated_sheet = ViewSheet.Create(doc, titleType if titleType else ElementId.InvalidElementId)
            duplicated_sheet.Name = sheet.Name
            duplicated_sheet.SetParamValue(SharedParamsConfig.Instance.AlbumBlueprints, self.sheet_album)
            if self.copyViews:
                for viewport in viewports:
                    viewport_type_id = viewport.GetTypeId()
                    if Viewport.CanAddViewToSheet(self.doc,
                                                  duplicated_sheet.Id,
                                                  viewport.ViewId):
                        duplicated_viewport = Viewport.Create(self.doc,
                                                              duplicated_sheet.Id,
                                                              viewport.ViewId,
                                                              viewport.GetBoxCenter())
                        duplicated_viewport.ChangeTypeId(viewport_type_id)

                    else:
                        original_view = doc.GetElement(viewport.ViewId)
                        view_name = original_view.Name
                        duplicated_view_id = original_view.Duplicate(ViewDuplicateOption.WithDetailing)
                        duplicated_view = doc.GetElement(duplicated_view_id)

                        duplicated_view_name = []
                        if self.prefix:
                            duplicated_view_name.append(self.prefix)

                        duplicated_view_name.append(view_name)

                        if self.suffix:
                            duplicated_view_name.append(self.suffix)

                        duplicated_view_name = "_".join(duplicated_view_name)

                        if len(duplicated_view_name) > len(view_name):
                            duplicated_view.Name = duplicated_view_name
                        try:
                            duplicated_view.SetParamValue(ProjectParamsConfig.Instance.ViewGroup, self.purpose)
                        except:
                            duplicated_view.ViewTemplateId = ElementId(-1)
                            duplicated_view.SetParamValue(ProjectParamsConfig.Instance.ViewGroup, self.purpose)

                        duplicated_viewport = Viewport.Create(self.doc,
                                                              duplicated_sheet.Id,
                                                              duplicated_view_id,
                                                              viewport.GetBoxCenter())
                        duplicated_viewport.ChangeTypeId(viewport_type_id)

            if details:
                ElementTransformUtils.CopyElements(sheet, details, duplicated_sheet, None, None)
            if textNotes:
                ElementTransformUtils.CopyElements(sheet, textNotes, duplicated_sheet, None, None)
            if scheduleSheets:
                ElementTransformUtils.CopyElements(sheet, scheduleSheets, duplicated_sheet, None, None)

        return duplicated_sheet

    def GetViewPurposeList(self):
        views = FilteredElementCollector(self.doc).OfCategory(BuiltInCategory.OST_Views).ToElements()
        purpose = []
        for view in views:
            param = view.GetParamValueOrDefault(ProjectParamsConfig.Instance.ViewGroup)
            if param and param not in purpose:
                purpose.append(param)
        return purpose

    def GetSheetAlbumList(self):
        sheets = FilteredElementCollector(doc) \
            .OfClass(ViewSheet) \
            .ToElements()
        sheet_album_list = []
        for sh in sheets:
            if sh.GetParamValueOrDefault(SharedParamsConfig.Instance.AlbumBlueprints):
                param = sh.GetParamValueOrDefault(SharedParamsConfig.Instance.AlbumBlueprints)
                if param not in sheet_album_list:
                    sheet_album_list.append(param)
        return sheet_album_list

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
