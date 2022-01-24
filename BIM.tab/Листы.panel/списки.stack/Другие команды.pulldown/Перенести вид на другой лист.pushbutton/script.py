# -*- coding: utf-8 -*-

import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script
from Autodesk.Revit.DB import *

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone.SharedParams import *

import re
from System.Windows.Data import CollectionViewSource, PropertyGroupDescription
from System.Windows.Media import VisualTreeHelper

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application


class SheetOption(object):
    def __init__(self, sheet, state=False):
        self.state = state
        self.name = '{}-{}'.format(sheet.SheetNumber, sheet.Name)
        self.Id = sheet.Id
        self.Number = sheet.SheetNumber
        self.sheet_album = sheet.GetParamValueOrDefault(SharedParamsConfig.Instance.AlbumBlueprints)
        self.GetAllViewports = sheet.GetAllViewports()
        self.foreGround = "#FF32329C"

        def __nonzero__(self):
            return self.state

        def __str__(self):
            return self.name


def list_albums(lists):
    albums = []
    for li in lists:
        if not li in albums:
            albums.append(li)
    return albums


def GroupByParameter(lst, func):
    res = {}
    for el in lst:
        key = func(el)
        if key in res:
            res[key].append(el)
        else:
            res[key] = [el]
    return res


def sort_fun(str):
    num = [int(x) for x in re.findall(r'\d+', str)]
    if len(num) > 0:
        return num[0]
    else:
        return 1000


class PrintSheetsWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, **kwargs):
        forms.WPFWindow.__init__(self, xaml_file_name)

        self.sheet_list = kwargs.get('list', None)

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
        border = VisualTreeHelper.GetChild(self.sheets_lb, 0)
        self.scrollViewer = border.Child
        offset = self.scrollViewer.VerticalOffset
        self._set_states(state=True, selected=True)
        self.scrollViewer.ScrollToVerticalOffset(offset)

    def uncheck_selected(self, sender, args):
        """Mark selected checkboxes as unchecked."""
        border = VisualTreeHelper.GetChild(self.sheets_lb, 0)
        self.scrollViewer = border.Child
        offset = self.scrollViewer.VerticalOffset
        self._set_states(state=False, selected=True)
        self.scrollViewer.ScrollToVerticalOffset(offset)

    def button_select(self, sender, args):
        """Handle select button click."""
        try:
            sheets = [x for x in self.sheet_list if x.state]
            if len(sheets) > 0:
                self.response = {
                    'sheets': sheets}
                self.Close()
            else:
                forms.alert('Вы должны выбрать хотя бы один лист для добавления '
                            'выбранные виды.')
        except:
            script.exit()


def getSheets():
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
    groupDescription = PropertyGroupDescription('sheet_album')
    view.GroupDescriptions.Add(groupDescription)
    window = PrintSheetsWindow('SelectsLists.xaml', list=view)

    window.ShowDialog()
    if hasattr(window, 'response'):
        res = window.response
        sheets = window.response
        selSheets = sheets['sheets']
        return selSheets


selViewports = []

cursheet = revit.active_view
if not isinstance(cursheet, DB.ViewSheet):
    forms.alert('Откройте лист, с которого надо перенести виды.')
    script.exit()
sel = revit.pick_elements()
if sel:
    for el in sel:
        selViewports.append(el)

if len(selViewports) > 0:
    dest_sheet = getSheets()
    for sh in dest_sheet:
        with revit.Transaction('Move Viewports'):
            for vp in selViewports:
                if isinstance(vp, DB.Viewport):
                    viewId = vp.ViewId
                    vpCenter = vp.GetBoxCenter()
                    vpTypeId = vp.GetTypeId()
                    cursheet.DeleteViewport(vp)
                    nvp = DB.Viewport.Create(revit.doc,
                                             sh.Id,
                                             viewId,
                                             vpCenter)
                    nvp.ChangeTypeId(vpTypeId)
                elif isinstance(vp, DB.ScheduleSheetInstance):
                    nvp = \
                        DB.ScheduleSheetInstance.Create(
                            revit.doc, sh.Id, vp.ScheduleId, vp.Point
                        )
                    revit.doc.Delete(vp.Id)
