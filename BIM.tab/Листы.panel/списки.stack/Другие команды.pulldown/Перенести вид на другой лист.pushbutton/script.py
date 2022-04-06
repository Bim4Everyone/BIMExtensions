# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from pyrevit import revit, DB
from pyrevit import forms
from pyrevit import script
from pyrevit import EXEC_PARAMS

from Autodesk.Revit.DB import *

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone.SharedParams import *

from dosymep_libs.bim4everyone import *

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
            if len(sheets) == 1:
                self.response = {'sheets': sheets}
                self.Close()
            else:
                forms.alert('Вы должны выбрать один лист для переноса выбранных видов.')
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
            if z.Id != revit.active_view.Id:
                sorted_sheets.append(z)

    view = CollectionViewSource.GetDefaultView(sorted_sheets)
    group_description = PropertyGroupDescription('sheet_album')
    view.GroupDescriptions.Add(group_description)
    window = PrintSheetsWindow('SelectsLists.xaml', list=view)

    window.ShowDialog()
    if hasattr(window, 'response'):
        res = window.response
        sheets = window.response
        sel_sheets = sheets['sheets']
        return sel_sheets


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    cur_sheet = revit.active_view
    if not isinstance(cur_sheet, DB.ViewSheet):
        forms.alert('Откройте лист, с которого надо перенести виды.', exitscript=True)

    sel_viewports = revit.pick_elements()
    if not sel_viewports:
        forms.alert('Не были выбраны виды для переноса.', exitscript=True)

    sel_viewports = [view_ports for view_ports in sel_viewports if
                     isinstance(view_ports, DB.Viewport) or isinstance(view_ports, DB.ScheduleSheetInstance)]
    if not sel_viewports:
        forms.alert('Не были выбраны виды для переноса.', exitscript=True)

    dest_sheet = getSheets()
    if not dest_sheet:
        script.exit()

    with revit.Transaction('BIM: Перенос видов'):
        for sheet in dest_sheet:
            for view_port in sel_viewports:
                if isinstance(view_port, DB.Viewport):
                    view_id = view_port.ViewId
                    vp_center = view_port.GetBoxCenter()
                    vp_type_id = view_port.GetTypeId()

                    cur_sheet.DeleteViewport(view_port)
                    nvp = DB.Viewport.Create(revit.doc,
                                             sheet.Id,
                                             view_id,
                                             vp_center)
                    nvp.ChangeTypeId(vp_type_id)
                elif isinstance(view_port, DB.ScheduleSheetInstance):
                    nvp = \
                        DB.ScheduleSheetInstance.Create(
                            revit.doc, sheet.Id, view_port.ScheduleId, view_port.Point
                        )
                    revit.doc.Delete(view_port.Id)


script_execute()