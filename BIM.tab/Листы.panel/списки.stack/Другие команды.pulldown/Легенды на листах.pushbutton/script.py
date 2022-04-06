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

logger = script.get_logger()


def is_placable(view):
    if view and view.ViewType and view.ViewType in [DB.ViewType.Schedule,
                                                    DB.ViewType.DraftingView,
                                                    DB.ViewType.Legend,
                                                    DB.ViewType.CostReport,
                                                    DB.ViewType.LoadsReport,
                                                    DB.ViewType.ColumnSchedule,
                                                    DB.ViewType.PanelSchedule]:
        return True
    return False


def update_if_placed(vport, exst_vps):
    for exst_vp in exst_vps:
        if vport.ViewId == exst_vp.ViewId:
            exst_vp.SetBoxCenter(vport.GetBoxCenter())
            exst_vp.ChangeTypeId(vport.GetTypeId())
            return True
    return False


class SheetOption(object):
    def __init__(self, sheet, state=False):
        self.state = state
        self.name = '{}-{}'.format(sheet.SheetNumber, sheet.Name)
        self.Id = sheet.Id
        self.Number = sheet.SheetNumber
        self.sheet_album = sheet.GetParamValueOrDefault(SharedParamsConfig.Instance.AlbumBlueprints)
        self.sheet_album = self.sheet_album if self.sheet_album else "Без имени"

        self.GetAllViewports = sheet.GetAllViewports()
        self.foreGround = "#FF32329C"

        # self.Visib = visibility
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
                            'выбранные легенды.')
        except:
            script.exit()


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    cursheet = revit.active_view
    if not isinstance(cursheet, DB.ViewSheet):
        forms.alert('Откройте лист, с которого надо перенести легенды.', exitscript=True)

    selected_vps = revit.pick_elements()
    if not selected_vps:
        forms.alert('Хотя бы одна легенда должна быть выбрана.', exitscript=True)

    allSheetedSchedules = DB.FilteredElementCollector(revit.doc) \
        .OfClass(DB.ScheduleSheetInstance) \
        .ToElements()
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
    selSheets = ''
    if hasattr(window, 'response'):
        sheets = window.response
        selSheets = sheets['sheets']

    # get a list of viewports to be copied, updated
    if selSheets and len(selSheets) > 0:
        for v in selSheets:
            if cursheet.Id == v.Id:
                selSheets.remove(v)
        with revit.Transaction('Copy Viewports to Sheets'):
            for sht in selSheets:
                existing_vps = [revit.doc.GetElement(x)
                                for x in sht.GetAllViewports]
                existing_schedules = [x for x in allSheetedSchedules
                                      if x.OwnerViewId == sht.Id]
                for vp in selected_vps:
                    if isinstance(vp, DB.Viewport):
                        src_view = revit.doc.GetElement(vp.ViewId)
                        # check if viewport already exists
                        # and update location and type
                        if update_if_placed(vp, existing_vps):
                            continue
                        # if not, create a new viewport
                        elif is_placable(src_view):
                            new_vp = \
                                DB.Viewport.Create(revit.doc,
                                                   sht.Id,
                                                   vp.ViewId,
                                                   vp.GetBoxCenter())

                            new_vp.ChangeTypeId(vp.GetTypeId())
                        else:
                            logger.warning('Skipping {}. This view type '
                                           'can not be placed on '
                                           'multiple sheets.'
                                           .format(src_view.ViewName))
                    elif isinstance(vp, DB.ScheduleSheetInstance):
                        # check if schedule already exists
                        # and update location
                        for exist_sched in existing_schedules:
                            if vp.ScheduleId == exist_sched.ScheduleId:
                                exist_sched.Point = vp.Point
                                break
                        # if not, place the schedule
                        else:
                            DB.ScheduleSheetInstance.Create(revit.doc,
                                                            sht.Id,
                                                            vp.ScheduleId,
                                                            vp.Point)


script_execute()