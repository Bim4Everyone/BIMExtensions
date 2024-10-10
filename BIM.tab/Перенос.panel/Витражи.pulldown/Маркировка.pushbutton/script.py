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
            self.__view_model.error_text = "Введите уникальный идентификатор"
            return False
        if not start_number_for_marking.isdecimal():
            self.__view_model.error_text = "Начальное значение некорректно"
            return False

        self.__view_model.error_text = None
        return True

    def Execute(self, parameter):
        # Выбираем элементы в модели
        selected_ids = uidoc.Selection.GetElementIds()

        # Отделяем только по одному экземпляру из всех выделенных групп
        all_groups = self.group_instance(selected_ids)

        with revit.Transaction("BIM: Маркировка"):
            list_groups = []
            for gro in all_groups:
                element_type = doc.GetElement(gro.GetTypeId())
                list_groups.append([element_type, self.get_area(gro)])

            sorted_list_group = sorted(list_groups, key=itemgetter(1))
            prefix = self.__view_model.marking_prefix
            start = self.__view_model.start_number
            suffix = self.__view_model.marking_suffix
            self.numeration(sorted_list_group, start, prefix, suffix)

        alert("Выполнение скрипта завершено")
        return True

    def location_up(self, element):
        point_down = element.Location.Point.Z * 304.8
        height = element.LookupParameter('ФОП_РАЗМ_Высота проёма').AsDouble() * 304.8
        point_up = point_down + height

        return round(point_up)

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

    def get_area(self, group):
        main_members = []
        for gr_member in self.group_members(group):
            if "Окн_Вит_Витраж" == gr_member.Name:
                main_members.append(gr_member)

        groups = {}
        for element in main_members:

            if str(self.location_up(element)) in groups:
                groups[str(self.location_up(element))].append(element)
            else:
                groups[str(self.location_up(element))] = [element]

        list_mains_mont = []
        list_corners_mont = []
        list_height_mont = []
        # Цикл по сгруппированным модулям по каждой высоте
        for items in groups.items():

            height_mont = []
            mains_width_mont = []
            left_corners_width_mont = []
            right_corners_width_mont = []
            # Цикл внутри каждой высоты
            for item in items[1]:

                if 0 == item.LookupParameter('Левый_Угловой').AsInteger() and 0 == item.LookupParameter(
                        'Правый_Угловой').AsInteger():
                    mains_width_mont.append(item.LookupParameter('ФОП_РАЗМ_Ширина проёма').AsDouble() * 304.8)
                    height_mont.append(item.LookupParameter('ФОП_РАЗМ_Высота проёма').AsDouble() * 304.8)

                elif 1 == item.LookupParameter('Левый_Угловой').AsInteger():
                    left_corners_width_mont.append(item.LookupParameter('ФОП_РАЗМ_Ширина проёма').AsDouble() * 304.8)
                    height_mont.append(item.LookupParameter('ФОП_РАЗМ_Высота проёма').AsDouble() * 304.8)

                elif 1 == item.LookupParameter('Правый_Угловой').AsInteger():
                    right_corners_width_mont.append(item.LookupParameter('ФОП_РАЗМ_Ширина проёма').AsDouble() * 304.8)
                    height_mont.append(item.LookupParameter('ФОП_РАЗМ_Высота проёма').AsDouble() * 304.8)

            sum_width_main_mont = sum(mains_width_mont)

            sum_width_left_corners_mont = sum(left_corners_width_mont)
            sum_width_right_corners_mont = sum(right_corners_width_mont)

            list_mains_mont.append(sum_width_main_mont)
            list_corners_mont.append(sum_width_left_corners_mont)
            list_corners_mont.append(sum_width_right_corners_mont)

            list_height_mont.append(max(height_mont))

        max_mains_mont = max(list_mains_mont)
        max_corners_mont = max(list_corners_mont)
        sum_height_mont = sum(list_height_mont)

        if round(max_corners_mont) == 0:
            corners_mont = ''
        else:
            corners_mont = 'x' + str(int(round(max_corners_mont)))

        mains_mont = str(int(round(max_mains_mont)))
        height_mont = 'x' + str(int(round(sum_height_mont))) + '(h)'

        gabarit_mont = mains_mont + height_mont + corners_mont

        list_gab = gabarit_mont.split('x')

        list_int_gab = []
        for gab in list_gab:
            list_int_gab.append(float(gab.replace('(h)', '')))

        if len(list_int_gab) == 2:
            area = (list_int_gab[0] / 1000) * (list_int_gab[1] / 1000)
        else:
            area = ((list_int_gab[0] / 1000) * (list_int_gab[1] / 1000)) + (
                        (list_int_gab[1] / 1000) * (list_int_gab[2] / 1000))

        return area

    # функция маркировки элементов
    def numeration(self, list, start, prefix, suffix):
        for count, element in enumerate(list):
            num = str(count) + start
            name = prefix + str(num) + suffix

            try:
                element[0].LookupParameter('ФОП_Марка типа').Set(name)
            except:
                element[0].LookupParameter('ФОП_Марка Типа').Set(name)


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
        self.__marking_prefix = "ВНХ-"
        self.__start_number = "1"
        self.__marking_suffix = ""
        self.__error_text = ""

        self.__create_command = CreateCommand(self)

    @reactive
    def marking_prefix(self):
        return self.__marking_prefix

    @marking_prefix.setter
    def marking_prefix(self, value):
        self.__marking_prefix = value

    @reactive
    def start_number(self):
        return self.__start_number

    @start_number.setter
    def start_number(self, value):
        self.__start_number = value

    @reactive
    def marking_suffix(self):
        return self.__marking_suffix

    @marking_suffix.setter
    def marking_suffix(self, value):
        self.__marking_suffix = value

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