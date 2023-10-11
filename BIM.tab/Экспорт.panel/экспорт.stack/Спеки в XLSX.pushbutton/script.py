# -*- coding: utf-8 -*-
import os

import clr

clr.AddReference('System')
clr.AddReference('System.IO')
clr.AddReference("PresentationCore")
clr.AddReference("PresentationFramework")

clr.AddReference("EPPlus.dll")
clr.AddReference("Microsoft.IO.RecyclableMemoryStream.dll")

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from System.IO import FileInfo
from System.Windows.Media import VisualTreeHelper
from System.Windows.Data import CollectionViewSource, PropertyGroupDescription

from OfficeOpenXml import *
from Autodesk.Revit.DB import *

from dosymep_libs.bim4everyone import *

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig

from pyrevit import forms
from pyrevit import script
from pyrevit import EXEC_PARAMS


def FilterString(obj):
    res = obj
    unacceptable_symbols = ['/', '\\', ':', '*', '<', '>', '|']
    for symbol in unacceptable_symbols:
        temp = res.split(symbol)
        res = "".join(temp)

    return res


class Excel(object):
    def __init__(self, obj, name):
        self.excel = obj
        self.name = name


class TabelsConverter(object):
    def __init__(self, views, save2_1file):
        excel_files = []
        excel_file = Excel(self.__create_package(), views[0].Name)
        for v in views:
            self.__currentRow = 1
            self.__currentColumn = 1
            data = v.GetTableData()
            if not save2_1file:
                excel_file = Excel(self.__create_package(), v.Name)
                excel_files.append(excel_file)

            current_worksheet = excel_file.excel.Workbook.Worksheets.Add(FilterString(v.Name))
            self.__export_to_excel(data, current_worksheet, v)

        if save2_1file:
            self.excel_file = [excel_file]
        else:
            self.excel_file = excel_files

    def __create_package(self):
        return ExcelPackage()

    def __export_to_excel(self, data, current_worksheet, v):
        self.__align_cells(data, current_worksheet)
        self.__export_section(SectionType.Header, data, current_worksheet, v)
        self.__export_section(SectionType.Body, data, current_worksheet, v)

    def __align_cells(self, data, current_worksheet):
        section_data = data.GetSectionData(SectionType.Body)
        number_of_columns = section_data.NumberOfColumns
        for columnNumber in range(number_of_columns):
            column_width = section_data.GetColumnWidthInPixels(columnNumber) / 6
            current_worksheet.Column(columnNumber + 1).Width = column_width

    def __export_section(self, section_type, data, current_worksheet, v):
        section_data = data.GetSectionData(section_type)

        number_of_rows = section_data.NumberOfRows
        number_of_columns = section_data.NumberOfColumns
        first_row = section_data.FirstRowNumber
        first_column = section_data.FirstColumnNumber

        for row_number in range(first_row, first_row + number_of_rows):
            row_height = section_data.GetRowHeightInPixels(row_number)
            current_worksheet.Row(self.__currentRow).Height = row_height
            for columnNumber in range(first_column, first_column + number_of_columns):
                self.__export_cell(row_number, columnNumber, section_type, data, current_worksheet, v)
                self.__currentColumn += 1
            self.__currentRow += 1
            self.__currentColumn = 1

    def __export_cell(self, row, column, section_type, data, current_worksheet, v):
        section_data = data.GetSectionData(section_type)
        merget_cell = section_data.GetMergedCell(row, column)
        if row == merget_cell.Top and column == merget_cell.Left:
            text = v.GetCellText(section_type, row, column)

            cell_style = section_data.GetTableCellStyle(row, column)

            excel_cell = current_worksheet.Cells[self.__currentRow, self.__currentColumn]

            excel_cell.Value = text

            split_text = text.split(",")
            if len(split_text) < 3:
                # print "may digit"
                if all(map(lambda x: x.isdigit(), split_text)):
                    excel_cell.Value = float(".".join(split_text))
                    if len(split_text) == 1:
                        excel_cell.Style.Numberformat.Format = "0"
                    else:
                        excel_cell.Style.Numberformat.Format = "0." + "0" * len(split_text[1])
                # print "DIGIT!!"

            excel_cell.Style.Font.Bold = cell_style.IsFontBold
            excel_cell.Style.Font.Italic = cell_style.IsFontItalic
            excel_cell.Style.Font.UnderLine = cell_style.IsFontUnderline
            excel_cell.Style.Font.Size = cell_style.TextSize
            excel_cell.Style.Font.Name = cell_style.FontName
            excel_cell.Style.WrapText = True

            horizontal_alignment = cell_style.FontHorizontalAlignment
            if horizontal_alignment == HorizontalAlignmentStyle.Center:
                excel_cell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
            elif horizontal_alignment == HorizontalAlignmentStyle.Left:
                excel_cell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
            else:
                excel_cell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

            vertical_alignment = cell_style.FontVerticalAlignment
            if vertical_alignment == VerticalAlignmentStyle.Top:
                excel_cell.Style.VerticalAlignment = Style.ExcelVerticalAlignment.Top
            elif vertical_alignment == VerticalAlignmentStyle.Middle:
                excel_cell.Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
            else:
                excel_cell.Style.VerticalAlignment = Style.ExcelVerticalAlignment.Bottom

            row_range = merget_cell.Bottom - row
            column_range = merget_cell.Right - column
            if row_range > 0 or column_range > 0:
                current_worksheet.Cells[self.__currentRow,
                self.__currentColumn, self.__currentRow + row_range, self.__currentColumn + column_range].Merge = True


class CheckBoxOption:
    def __init__(self, obj):
        self.state = False
        self.name = obj.Name
        self.obj = obj
        self.view_assignment = obj.GetParamValueOrDefault(ProjectParamsConfig.Instance.ViewGroup)
        self.view_assignment = self.view_assignment if self.view_assignment else None
        self.is_open = "Все спецификации"
        self.sort_order = 3

    def __nonzero__(self):
        return self.state

    def __str__(self):
        return self.name


class PrintSheetsWindow(forms.WPFWindow):
    def __init__(self, xaml_file_name, **kwargs):
        forms.WPFWindow.__init__(self, xaml_file_name)

        self.sheet_list = kwargs.get('list', None)
        self.active_view = kwargs.get('active_view', None)
        self.select_active_view()

    def select_active_view(self):
        if self.active_view:
            for element in self.sheets_lb.ItemsSource:
                if element.name == self.active_view.Name:
                    element.state = True

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

    def _set_states(self, state=True, flip=False, selected=False):
        all_items = self.sheet_list

        if selected:
            current_list = self.selected_sheets
        else:
            current_list = self.sheet_list
        for checkbox in current_list:
            if flip:
                checkbox.state = not checkbox.state
            else:
                checkbox.state = state

        # push list view to redraw
        self.sheet_list = []
        self.sheet_list = all_items

    def toggle_all(self, sender, args):
        """Handle toggle all button to toggle state of all check boxes."""
        self._set_states(flip=True)

    def check_all(self, sender, args):
        """Handle check all button to mark all check boxes as checked."""
        self._set_states(state=True)

    def uncheck_all(self, sender, args):
        """Handle uncheck all button to mark all check boxes as un-checked."""
        self._set_states(state=False)

    def check_selected(self, sender, args):
        """Mark selected checkboxes as checked."""
        border = VisualTreeHelper.GetChild(self.sheets_lb, 0)
        self.scroll_viewer = border.Child
        offset = self.scroll_viewer.VerticalOffset
        self._set_states(state=True, selected=True)
        self.scroll_viewer.ScrollToVerticalOffset(offset)

    def uncheck_selected(self, sender, args):
        """Mark selected checkboxes as unchecked."""
        border = VisualTreeHelper.GetChild(self.sheets_lb, 0)
        self.scroll_viewer = border.Child
        offset = self.scroll_viewer.VerticalOffset
        self._set_states(state=False, selected=True)
        self.scroll_viewer.ScrollToVerticalOffset(offset)

    def button_select(self, sender, args):
        """Handle select button click."""
        try:
            sheets = [x for x in self.sheet_list if x.state]
            if len(sheets) > 0:
                self.response = {
                    'sheets': sheets,
                    'oneFile': self.oneFile.IsChecked}
                self.Close()
            else:
                forms.alert('Вы должны выбрать хотя бы один лист.')
        except:
            script.exit()


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    app = __revit__.Application
    view = __revit__.ActiveUIDocument.ActiveGraphicalView
    active_view = doc.ActiveView

    opened_views = [x.ViewId for x in uidoc.GetOpenUIViews()]

    checked_view = None
    if active_view.ViewType == ViewType.Schedule:
        checked_view = active_view

    viewSchedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
    items = [CheckBoxOption(x) for x in viewSchedules]

    for item in items:
        if item.obj.Id in opened_views:
            item.is_open = "Открытые спецификации"
            item.sort_order = 2
        if item.obj.Id == view.Id:
            item.is_open = "Активная спецификация"
            item.sort_order = 1

    items = sorted(items, key=lambda item: (item.sort_order, item.name))

    view = CollectionViewSource.GetDefaultView(items)
    groupDescription = PropertyGroupDescription('is_open')
    view.GroupDescriptions.Add(groupDescription)

    window = PrintSheetsWindow('selectViews.xaml', list=view, active_view=checked_view)
    window.ShowDialog()

    sel_sheets = ''
    if hasattr(window, 'response'):
        sheets = window.response
        sel_sheets = sheets['sheets']
        save2_1file = sheets['oneFile']

    if not sel_sheets:
        script.exit()

    folder_name = forms.pick_folder(title="Выберите папку для сохранения спецификаций")
    if folder_name:
        views = []
        file_info = ''

        for item in sel_sheets:
            if file_info == '':
                file_info = FileInfo(folder_name + "\\" + FilterString(item.name) + ".xlsx")
            views.append(item.obj)

        converter = TabelsConverter(views, save2_1file)
        package = converter.excel_file
        try:
            count = 1
            for p in package:
                path_file = folder_name + "\\" + FilterString(p.name) + ".xlsx"
                count = 1
                while os.path.isfile(path_file):
                    path_file = folder_name + "\\" + FilterString(p.name) + "-" + str(count) + ".xlsx"
                    count += 1

                file_info = FileInfo(path_file)
                p.excel.SaveAs(file_info)
                p.excel.Save()
                p.excel.Dispose()

            show_executed_script_notification()
        except Exception as ex:
            print "Не удалось сохранить спецификацию '{}'".format(item.name)
            print "Исключение '{}'".format(ex)


script_execute()
