# -*- coding: utf-8 -*-
import clr
import datetime

from System.Collections.Generic import *
from operator import itemgetter

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from System.Windows.Input import ICommand

import pyevent
from pyrevit import EXEC_PARAMS, revit
from pyrevit.forms import *
from pyrevit import script

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application


IGNORE_WALLS = ['теплитель',
			'КЖ',
			'кж',
			'(КР)',
			'(К)',
			'(K)',
			'Железобетон',
			'железобетон',
			'ЖБ',
			'жб',
			'Монолит',
			'монолит',
			'Маячк',
			'маячк',
			'ПРФЛ',
			'1 блок']

class CreateCommand(ICommand):
    CanExecuteChanged, _canExecuteChanged = pyevent.make_event()

    def __init__(self, view_model, *args):
        ICommand.__init__(self, *args)
        self.__view_model = view_model
        self.__view_model.PropertyChanged += self.ViewModel_PropertyChanged

    def add_CanExecuteChanged(self, value):
        self.CanExecuteChanged += value

    def remove_CanExecuteChanged(self, value):
        self.CanExecuteChanged -= value

    def OnCanExecuteChanged(self):
        # В Python при работе с событиями нужно явно
        # передавать импорт в обработчике события
        from System import EventArgs
        self._canExecuteChanged(self, EventArgs.Empty)

    def ViewModel_PropertyChanged(self, sender, e):
        self.OnCanExecuteChanged()

    def CanExecute(self, parameter):
        start_number_for_marking = self.__view_model.start_number
        if start_number_for_marking == "":
            self.__view_model.error_text = "Введите начальное значение"
            return False
        if not start_number_for_marking.isdecimal():
            self.__view_model.error_text = "Начальное значение некорректно"
            return False

        self.__view_model.error_text = None
        return True

    def Execute(self, parameter):
        SELECTED_IDS = uidoc.Selection.GetElementIds()
        self.analize_fam(SELECTED_IDS)
        all_elements = [doc.GetElement(id) for id in SELECTED_IDS]

        if len(all_elements) == 0:
            alert("Ничего не выбрано. Выберите элементы.")

        with revit.Transaction("BIM: Подготовка к маркировке"):
            result_openings_list = []
            for element in all_elements:
                host = element.Host.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()
                if not any(IGNORE_WALL in host for IGNORE_WALL in IGNORE_WALLS):
                    result_openings_list.append(element)
                else:
                    element.LookupParameter('ФОП_Марка проема').Set('')

        with revit.Transaction("BIM: Маркировка"):
            groupped_sorted_list = self.group_elements_list(result_openings_list)
            elements = []
            errors = []

            # маркируем все группы элементов
            for count, element_group in enumerate(groupped_sorted_list):
                marka_element = self.__view_model.prefix + str(count + int(self.__view_model.start_number))
                for element in element_group:
                    try:
                        elements.append(element.LookupParameter('ФОП_Марка проема').Set(marka_element))
                    except:
                        errors.append(element.Id)

        alert("Выполнение скрипта завершено")
        return True

    def analize_fam(self, list_id):
        for id in list_id:
            fam = doc.GetElement(id).Category.Name
            if fam == 'Окна' or fam == 'Двери':
                pass
            else:
                alert("Выбраны недопустимые элементы.")

    def group_elements_list(self, list):
        # создаем словарь и группируем отсортированный список по коду
        groups = {}
        for element in list:
            groupby_value = element.LookupParameter('ФОП_Группирование').AsString()
            if groupby_value in groups:
                groups[groupby_value].append(element)
            else:
                groups[groupby_value] = [element]

        # возвращаем список сгруппированных элементов
        groupped_list = groups.Values

        # создаем список и вытаскиваем параметр ширины для сортировки
        list_element = []
        for el in groupped_list:
            len = int(el[0].LookupParameter('ФОП_РАЗМ_Ширина проёма').AsValueString())
            list_element.append([el, len])

        # сортируем сгруппированный список и создаем отсортированный список
        sorted_list_element = sorted(list_element, key=itemgetter(1))

        # убираем из списка ключ сортировки
        result_list = []
        for sorted_element in sorted_list_element:
            result_list.append(sorted_element[0])

        return result_list


class MainWindow(WPFWindow):
    def __init__(self):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), "MainWindow.xaml")
        super(MainWindow, self).__init__(self.xaml_source)

    def ButtonOK_Click(self, sender, e):
        self.DialogResult = True

    def ButtonCancel_Click(self, sender, e):
        self.DialogResult = False
        self.Close()


class MainWindowViewModel(Reactive):
    def __init__(self):
        Reactive.__init__(self)
        self.__prefix = ""
        self.__start_number = "1"
        self.__error_text = ""

        self.__create_command = CreateCommand(self)

    @reactive
    def prefix(self):
        return self.__prefix

    @prefix.setter
    def prefix(self, value):
        self.__prefix = value

    @reactive
    def start_number(self):
        return self.__start_number

    @start_number.setter
    def start_number(self, value):
        self.__start_number = value

    @reactive
    def error_text(self):
        return self.__error_text

    @error_text.setter
    def error_text(self, value):
        self.__error_text = value

    @property
    def create_command(self):
        return self.__create_command


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel()
    main_window.show_dialog()
    if not main_window.DialogResult:
        script.exit()


script_execute()