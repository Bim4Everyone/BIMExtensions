# -*- coding: utf-8 -*-
import clr
import datetime

from System.Collections.Generic import *
from operator import itemgetter

import itertools
from itertools import groupby

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
        return True

    def Execute(self, parameter):
        # Выбираем элементы в модели
        selected_ids = uidoc.Selection.GetElementIds()

        # Отделяем только по одному экземпляру из всех выделенных групп
        all_groups = self.group_instance(selected_ids)

        with revit.Transaction("BIM: Описание"):
            for gro in all_groups:
                element_type = doc.GetElement(gro.GetTypeId())
                try:
                    element_type.LookupParameter('ФОП_Описание типа').Set(self.description(gro))
                except:
                    element_type.LookupParameter('ФОП_Описание Типа').Set(self.description(gro))

        alert("Выполнение скрипта завершено")
        return True

    # Функция, которая из всех выделенных групп, берет по одной
    def group_instance(self, list_id):
        all_elements = []
        for id in list_id:
            noname_element = doc.GetElement(id)
            if "Группы модели" == noname_element.Category.Name:
                all_elements.append(noname_element)

        groups = {}
        for group in all_elements:
            group_by_value = group.Name

            if group_by_value in groups:
                groups[group_by_value].append(group)
            else:
                groups[group_by_value] = [group]

        elements = []
        for element in groups.values():
            elements.append(element[0])

        return elements

    # Функция, которая возвращает членов группы
    def group_members(self, group):

        gr_members_ids = group.GetMemberIds()

        list_of_members = []
        for id in gr_members_ids:
            member = doc.GetElement(id)
            list_of_members.append(member)

        return list_of_members

    def in_out(self, element):
        if 1 == element.LookupParameter('Наружный').AsInteger():
            return 'Наружный витраж'
        else:
            return 'Внутренний витраж'

    def warm_cold(self, element):
        if 1 == element.LookupParameter('Холодный').AsInteger():
            return ' из "холодного"'
        else:
            return ' из "тёплого"'

    def profile(self, element):
        if 0 == element.LookupParameter('Профиль').AsInteger():
            return ' алюминиевого профиля'
        else:
            return ' ПВХ-профиля'

    def glass(self, element):
        if 0 == element.LookupParameter('Остекление').AsInteger():
            return ' c одинарным остеклением.'
        elif 1 == element.LookupParameter('Остекление').AsInteger():
            return ' c однокамерным стеклопакетом.'
        elif 2 == element.LookupParameter('Остекление').AsInteger():
            return ' c двухкамерным стеклопакетом.'
        elif 3 == element.LookupParameter('Остекление').AsInteger():
            return ' c трехкамерным стеклопакетом.'
        else:
            return ' c двухкамерным стеклопакетом.'

    def ograjdenie(self, list):
        if sum(list) > 0:
            return ' Встроенное перильное ограждение, шаг 0,11м.'
        else:
            return ''

    def comparison_mat(self, name):
        if 'решетка' in name.lower():
            return ' Решетка в составе витража (c внутренней стороны закрыть временной сэндвич панелью).'
        elif 'тонированн' in name.lower():
            return ' Вставки из тонированного стекла.'
        elif 'закаленн' in name.lower():
            return ' Вставки из закалённого стекла.'
        elif 'сэндвич' in name.lower():
            return ' Вставки из сэндвич панелей.'
        else:
            return ''

    def parsing_mats(self, element):

        list = []

        if 'Окн_Вит_Ячейка_1х1' == element.Symbol.Family.Name:
            list.append(self.comparison_mat(element.Symbol.LookupParameter('Заполнение_Материал').AsValueString()))

        elif 'Окн_Вит_Ячейка_1х2' == element.Symbol.Family.Name:
            list.append(
                self.comparison_mat(element.Symbol.LookupParameter('Заполнение_Материал_Левая_Створка').AsValueString()))
            list.append(
                self.comparison_mat(element.Symbol.LookupParameter('Заполнение_Материал_Правая_Створка').AsValueString()))

        elif 'Окн_Вит_Ячейка_1х3' == element.Symbol.Family.Name:
            list.append(
                self.comparison_mat(element.Symbol.LookupParameter('Заполнение_Материал_Левая_Створка').AsValueString()))
            list.append(
                self.comparison_mat(element.Symbol.LookupParameter('Заполнение_Материал_Центр_Створка').AsValueString()))
            list.append(
                self.comparison_mat(element.Symbol.LookupParameter('Заполнение_Материал_Правая_Створка').AsValueString()))

        elif 'Окн_Вит_Ячейка_Дверь' == element.Symbol.Family.Name:
            list.append(self.comparison_mat(element.Symbol.LookupParameter('Заполнение_Материал').AsValueString()))
            if 1 == element.Symbol.LookupParameter(
                    'Домофонная_Панель_Справа').AsInteger() or 1 == element.Symbol.LookupParameter(
                    'Домофонная_Панель_Слева').AsInteger():
                list.append(self.comparison_mat(
                    element.Symbol.LookupParameter('Заполнение_Материал_Домофонная_Панель').AsValueString()))

        return list

    def open(self, element, par):
        return element.Symbol.LookupParameter(par).AsInteger()

    def parsing_cell(self, element):
        list = []

        if 'Окн_Вит_Ячейка_1х1' == element.Symbol.Family.Name:
            if 1 == self.open(element, 'Код_Открывания') or 2 == self.open(element, 'Код_Открывания'):
                list.append(' Поворотная створка.')
            elif 2 == self.open(element, 'Код_Открывания') or 3 == self.open(element, 'Код_Открывания'):
                list.append(' Поворотнo-откидная створка.')
            elif 5 == self.open(element, 'Код_Открывания') or 6 == self.open(element, 'Код_Открывания'):
                list.append(' Откидная створка.')
            elif 7 == self.open(element, 'Код_Открывания'):
                list.append(' Сдвижная створка.')

        if 'Окн_Вит_Ячейка_1х2' == element.Symbol.Family.Name:
            if (1 == self.open(element, 'Код_Открывания_Левая_Створка')
                    or 2 == self.open(element, 'Код_Открывания_Левая_Створка')):
                list.append(' Поворотная створка.')
            elif (1 == self.open(element, 'Код_Открывания_Правая_Створка')
                  or 2 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Поворотная створка.')

            elif (3 == self.open(element, 'Код_Открывания_Левая_Створка')
                  or 4 == self.open(element, 'Код_Открывания_Левая_Створка')):
                list.append(' Поворотнo-откидная створка.')
            elif (3 == self.open(element, 'Код_Открывания_Правая_Створка')
                  or 4 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Поворотнo-откидная створка.')

            elif (5 == self.open(element, 'Код_Открывания_Левая_Створка')
                  or 6 == self.open(element, 'Код_Открывания_Левая_Створка')):
                list.append(' Откидная створка.')
            elif (5 == self.open(element, 'Код_Открывания_Правая_Створка')
                  or 6 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Откидная створка.')

            elif 7 == self.open(element, 'Код_Открывания_Левая_Створка'):
                list.append(' Сдвижная створка.')
            elif 7 == self.open(element, 'Код_Открывания_Правая_Створка'):
                list.append(' Сдвижная створка.')

            elif 8 == self.open(element, 'Код_Открывания_Левая_Створка'):
                list.append(' Окно выдачи.')
            elif 8 == self.open(element, 'Код_Открывания_Правая_Створка'):
                list.append(' Окно выдачи.')

        if 'Окн_Вит_Ячейка_1х3' == element.Symbol.Family.Name:
            if (1 == self.open(element, 'Код_Открывания_Левая_Створка')
                    or 2 == self.open(element, 'Код_Открывания_Левая_Створка')):
                list.append(' Поворотная створка.')
            elif (1 == self.open(element, 'Код_Открывания_Центр_Створка')
                  or 2 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Поворотная створка.')
            elif (1 == self.open(element, 'Код_Открывания_Правая_Створка')
                  or 2 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Поворотная створка.')

            elif (3 == self.open(element, 'Код_Открывания_Левая_Створка')
                  or 4 == self.open(element, 'Код_Открывания_Левая_Створка')):
                list.append(' Поворотнo-откидная створка.')
            elif (3 == self.open(element, 'Код_Открывания_Центр_Створка')
                  or 4 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Поворотная створка.')
            elif (3 == self.open(element, 'Код_Открывания_Правая_Створка')
                  or 4 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Поворотнo-откидная створка.')

            elif (5 == self.open(element, 'Код_Открывания_Левая_Створка')
                  or 6 == self.open(element, 'Код_Открывания_Левая_Створка')):
                list.append(' Откидная створка.')
            elif (5 == self.open(element, 'Код_Открывания_Центр_Створка')
                  or 6 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Поворотная створка.')
            elif (5 == self.open(element, 'Код_Открывания_Правая_Створка')
                  or 6 == self.open(element, 'Код_Открывания_Правая_Створка')):
                list.append(' Откидная створка.')

            elif 7 == self.open(element, 'Код_Открывания_Левая_Створка'):
                list.append(' Сдвижная створка.')
            elif 7 == self.open(element, 'Код_Открывания_Центр_Створка'):
                list.append(' Сдвижная створка.')
            elif 7 == self.open(element, 'Код_Открывания_Правая_Створка'):
                list.append(' Сдвижная створка.')

        if 'Окн_Вит_Ячейка_Дверь' == element.Symbol.Family.Name:
            if (0 == self.open(element, 'Код_Открывания')
                    or 1 == self.open(element, 'Код_Открывания')):
                list.append(' Одностворчатая дверь.')
            elif (2 == self.open(element, 'Код_Открывания')
                  or 3 == self.open(element, 'Код_Открывания')
                  or 4 == self.open(element, 'Код_Открывания')):
                list.append(' Двустворчатая дверь (одна из створок не менее 900мм).')

            if 1 == self.open(element, 'Домофонная_Панель_Слева') or 1 == self.open(element, 'Домофонная_Панель_Справа'):
                list.append(' Домофонная панель.')

        return list

    # Функция, которая возвращает описание состава группы
    def description(self, group):
        members = []
        main_members = []
        visor = ''

        for gr_member in self.group_members(group):
            if "Окн_Вит_Витраж" != gr_member.Name:
                members.append(gr_member)
            elif "Окн_Вит_Витраж" == gr_member.Name:
                main_members.append(gr_member)

            if "ОбщМд_Козырек_Стеклянный" == gr_member.Name:
                visor = ' Козырек.'
            else:
                visor = ''

        # Начало описание по любому модулю
        main_code = (self.in_out(main_members[0]) + self.warm_cold(main_members[0])
                     + self.profile(main_members[0]) + self.glass(main_members[0]))

        # Описание ячеек и материалов
        mats = []
        cells = []
        for member in members:
            mats.append(self.parsing_mats(member))
            cells.append(self.parsing_cell(member))

        clean_mats = set(list(itertools.chain(*mats)))
        clean_cells = set(list(itertools.chain(*cells)))

        if len(clean_cells) == 0:
            cells_code = ' Глухой.'
        else:
            cells_code = ''.join(clean_cells)

        mats_code = ''.join(clean_mats)

        # Описание ограждения
        ogr_list = []
        for main in main_members:
            if int(main.LookupParameter('Ограждение_Высота').AsValueString()) > 0:
                ogr_list.append(1)

        ogr_code = self.ograjdenie(ogr_list)

        return main_code + cells_code + mats_code + ogr_code + visor + ' ' + self.__view_model.added_text


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
        self.__added_text = "Цвет профиля RAL7024 (матовый)."
        self.__error_text = ""

        self.__create_command = CreateCommand(self)

    @reactive
    def added_text(self):
        return self.__added_text

    @added_text.setter
    def added_text(self, value):
        self.__added_text = value

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