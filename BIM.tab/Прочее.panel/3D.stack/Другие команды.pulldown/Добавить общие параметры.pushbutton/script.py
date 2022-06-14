# coding=utf-8

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import pyevent #pylint: disable=import-error
from System.Windows.Input import ICommand

from dosymep_libs.bim4everyone import *

from pyrevit import *
from pyrevit.forms import *
from pyrevit.revit import *

from Autodesk.Revit.DB import *

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal


doc = __revit__.ActiveUIDocument.Document
uiapp = __revit__
app = uiapp.Application


class PickExcelFileCommand(ICommand):
    CanExecuteChanged, _canExecuteChanged = pyevent.make_event()

    def __init__(self, view_model, *args):
        ICommand.__init__(self, *args)
        self.__view_model = view_model

    def add_CanExecuteChanged(self, value):
        self.CanExecuteChanged += value

    def remove_CanExecuteChanged(self, value):
        self.CanExecuteChanged -= value

    def OnCanExecuteChanged(self):
        self._canExecuteChanged(self, System.EventArgs.Empty)

    def CanExecute(self, parameter):
        return True

    def Execute(self, parameter):
        self.__view_model.excel_path = pick_excel_file()


class PickTXTFileCommand(ICommand):
    CanExecuteChanged, _canExecuteChanged = pyevent.make_event()

    def __init__(self, view_model, *args):
        ICommand.__init__(self, *args)
        self.__view_model = view_model

    def add_CanExecuteChanged(self, value):
        self.CanExecuteChanged += value

    def remove_CanExecuteChanged(self, value):
        self.CanExecuteChanged -= value

    def OnCanExecuteChanged(self):
        self._canExecuteChanged(self, System.EventArgs.Empty)

    def CanExecute(self, parameter):
        return True

    def Execute(self, parameter):
        self.__view_model.txt_path = pick_file()


class MainWindowViewModel(Reactive):
    def __init__(self, *args):
        Reactive.__init__(self, *args)

        self.__excel_path = "Файл не выбран!"
        self.__txt_path = None
        self.__pick_excel_file_command = PickExcelFileCommand(self)
        self.__pick_txt_file_command = PickTXTFileCommand(self)


    @property
    def PickExcelFileCommand(self):
        return self.__pick_excel_file_command

    @reactive
    def excel_path(self):
        return self.__excel_path

    @excel_path.setter
    def excel_path(self, value):
        self.__excel_path = value

    @property
    def PickTXTFileCommand(self):
        return self.__pick_txt_file_command

    @reactive
    def txt_path(self):
        return self.__txt_path

    @txt_path.setter
    def txt_path(self, value):
        self.__txt_path = value


class MainWindow(WPFWindow):
    def __init__(self,):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')

        super(MainWindow, self).__init__(self.xaml_source)

    def open_excel_file(self, sender, e):
        self.excel_file = pick_excel_file()
        if self.excel_file:
            return self.excel_file

    def open_txt_file(self, sender, e):
        self.txt_file = pick_file(file_ext='txt')
        if self.txt_file:
            return self.txt_file

    def ButtonOK_Click(self, sender, e):
        self.run = True
        self.Close()


def read_from_excel(path, is_family):
    parameters = []
    categories = []
    excel = Excel.ApplicationClass()
    excel.Visible = False
    excel.DisplayAlerts = False

    try:
        workbook = excel.Workbooks.Open(path)

        ws_1 = workbook.Worksheets(1)
        row_end_1 = ws_1.Cells.Find("*", System.Reflection.Missing.Value,
                                    System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                    Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                    False, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row

        for i in range(2, row_end_1 + 1):
            parameter = []
            for j in range(1, 6):
                parameter.append(ws_1.Cells(i, j).Text)
            parameters.append(parameter)

        if not is_family:
            ws_2 = workbook.Worksheets(2)
            row_end_2 = ws_2.Cells.Find("*", System.Reflection.Missing.Value,
                                        System.Reflection.Missing.Value, System.Reflection.Missing.Value,
                                        Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious,
                                        False, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row

            for i in range(1, row_end_2 + 1):
                categories.append(ws_2.Cells(i, 1).Text)

    finally:
        excel.ActiveWorkbook.Close(False)
        Marshal.ReleaseComObject(ws_1)
        Marshal.ReleaseComObject(workbook)
        Marshal.ReleaseComObject(excel)
    return parameters, categories


def set_vary_by_group(doc, param_names, vary_by_group):
    b_map = doc.ParameterBindings
    iterator = b_map.ForwardIterator()
    result = True
    while iterator.MoveNext():
        int_def = iterator.Key
        if int_def.Name in param_names:
            with revit.Transaction("BIM: Настроено изменение параметров по группам"):
                try:
                    int_def.SetAllowVaryBetweenGroups(doc, vary_by_group)
                except Exception as exc:
                    result = exc
    return result


def get_category_by_name(doc, name):
    all_categories = doc.Settings.Categories
    for category in all_categories:
        if name == category.Name:
            return category
        sub_categories = category.SubCategories
        if sub_categories:
            for sub_category in sub_categories:
                if name == sub_category.Name:
                    if sub_category.AllowsBoundParameters:
                        return sub_category


def get_categories_list(doc, categories_name):
    categories = []
    errors = ""
    if categories_name:
        for category_name in categories_name:
            if category_name:
                category_revit = get_category_by_name(doc, category_name)
                if category_revit:
                    categories.append(category_revit)
                else:
                    errors += "Категория  с именем '{}' не найдена в проекте\n".format(category_name)
    if errors:
        errors = "Параметры не добавлены!\n" + errors
    return categories, errors


def create_ex_definition(param_group, param_name):
    shared_file = app.OpenSharedParameterFile()
    group_name = shared_file.Groups.get_Item(param_group)
    ex_definition = group_name.Definitions.get_Item(param_name)
    return ex_definition


def create_category_set(categories):
    category_set = app.Create.NewCategorySet()
    for category in categories:
        category_set.Insert(category)
    return category_set


def add_shared_parameter_to_project(doc, cat_set, ex_def, param_group, param_inst):
    if param_inst:
        new_bind = app.Create.NewInstanceBinding(cat_set)
    else:
        new_bind = app.Create.NewTypeBinding(cat_set)
    param_group = BuiltInParameterGroup.Parse(BuiltInParameterGroup, param_group)
    doc.ParameterBindings.Insert(ex_def, new_bind, param_group)
    return "Параметр '{}' добавлен\n".format(ex_def.Name)


def add_shared_parameter_to_family(doc, ex_def, param_group, param_inst):
    param_group = BuiltInParameterGroup.Parse(BuiltInParameterGroup, param_group)
    doc.FamilyManager.AddParameter(ex_def, param_group, param_inst)
    return "Параметр '{}' добавлен\n".format(ex_def.Name)


def add_category_to_shared_parameter(doc, cat_set, ex_def):
    bind_map = doc.ParameterBindings.Item[ex_def]
    for new_category in cat_set:
        bind_map.Categories.Insert(new_category)
        doc.ParameterBindings.ReInsert(ex_def, bind_map)
    return "К параметру '{}' добавлены категории\n".format(ex_def.Name)


def check_and_add_parameter(doc,
                            categories,
                            group_in_txt,
                            param_name,
                            param_group,
                            param_inst
                            ):
    with revit.Transaction("BIM: Добавлен общий параметр"):
        definitions = []
        ex_def = create_ex_definition(group_in_txt, param_name)
        definitions.append(ex_def)
        for ex_def in definitions:
            if doc.IsFamilyDocument:
                parameters_in_family = [param.Definition.Name for param in doc.FamilyManager.GetParameters()]
                if ex_def.Name not in parameters_in_family:
                    result = add_shared_parameter_to_family(doc, ex_def, param_group, param_inst)
                else:
                    result = "Параметр '{}' уже есть в семействе\n".format(ex_def.Name)
            else:
                cat_set = create_category_set(categories)
                if doc.ParameterBindings.Contains(ex_def):
                    result = add_category_to_shared_parameter(doc, cat_set, ex_def)
                else:
                    result = add_shared_parameter_to_project(doc, cat_set, ex_def, param_group, param_inst)
    return result


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    try:
        spf_path = app.SharedParametersFilename
    except:
        spf_path = "Файл не выбран!"

    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel()
    main_window.DataContext.txt_path = spf_path
    main_window.show_dialog()

    excel_path = main_window.DataContext.excel_path
    new_spf_path = main_window.DataContext.txt_path

    app.SharedParametersFilename = new_spf_path

    if excel_path and main_window.run:
        parameters, categories_str = read_from_excel(excel_path, doc.IsFamilyDocument)

        categories, error_message = get_categories_list(doc, categories_str)

        if error_message:
            print error_message
            script.exit()

        result_message = ""
        for parameter in parameters:
            param_name = parameter[0]
            group_in_txt = parameter[1]
            param_group = parameter[2]
            param_inst = int(parameter[3])

            result = check_and_add_parameter(doc,
                                             categories,
                                             group_in_txt,
                                             param_name,
                                             param_group,
                                             param_inst
                                             )

            if not doc.IsFamilyDocument:
                vary_by_group = int(parameter[4])
                set_vary_by_group(doc,
                                  param_name,
                                  vary_by_group
                                  )
            result_message += result
        print result_message


script_execute()
