# -*- coding: utf-8 -*-
import clr
import datetime

from System.Collections.Generic import *

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


KR_LIST = ['(КР)',
			'(К)',
			'(K)',
			'КЖ',
			'кж'
			'ЖБ',
			'жб',
			'Железобетон',
			'железобетон',
			'Монолит',
			'монолит',
			'Балка']


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
			'1 блок']

BASE_POINT = 0

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
        if self.__view_model.unique_identifier == "":
            self.__view_model.error_text = "Введите уникальный идентификатор"
            return False

        self.__view_model.error_text = None
        return True

    def Execute(self, parameter):
        VIEWS = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_Views)
                 .WhereElementIsNotElementType()
                 .ToElements())
        global BASE_POINT
        BASE_POINT = (FilteredElementCollector(doc)
                      .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                      .WhereElementIsNotElementType()
                      .ToElements())
        SELECTED_IDS = uidoc.Selection.GetElementIds()

        UNIQUE = self.__view_model.unique_identifier

        # Анализируем выделение на семейства окон и дверей
        self.analize_fam(SELECTED_IDS)

        # проверяем, выбрано хоть что-то
        all_elements = [doc.GetElement(id) for id in SELECTED_IDS]
        if len(all_elements) == 0:
            alert("Ничего не выбрано. Выберите элементы.")

        with revit.Transaction("BIM: Создаем 3D-вид"):
            # Создаем 3Д вид
            self.create_3D(VIEWS)

        with revit.Transaction("BIM: Получаем 3D-вид"):
            # Получаем 3Д вид
            VIEWS = FilteredElementCollector(doc).OfCategory(
                BuiltInCategory.OST_Views).WhereElementIsNotElementType().ToElements()
            list_of_3D = []

            for VIEW in VIEWS:
                if '3D Проемы и отверстия' == VIEW.Name:
                    list_of_3D.append(VIEW)
                    view_3D = list_of_3D[0]

        with revit.Transaction("BIM: Создаем 3D-вид"):
            result_openings_list = []
            for element in all_elements:
                self.set_parameter(element, UNIQUE, view_3D)


        alert("Выполнено успешно!")
        return True

    def analize_fam(self, list_id):
        for id in list_id:
            fam = doc.GetElement(id).Category.Name
            if fam == 'Окна' or fam == 'Двери':
                pass
            else:
                alert("Выбраны недопустимые элементы.")

    def create_3D(self, VIEWS):
        view_types = FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements()
        view_types_3D = [vt for vt in view_types if vt.ViewFamily == ViewFamily.ThreeDimensional]
        view_type_3D = view_types_3D[0]

        list_of_3D = []

        if '3D Проемы и отверстия' in [VIEW.Name for VIEW in VIEWS]:
            list_of_3D.append(VIEW)
            return list_of_3D[0]
        else:
            View3D.CreateIsometric(doc, view_type_3D.Id).Name = '3D Проемы и отверстия'

    def set_parameter(self, element, unique, view_3D):
        kr_proximity = self.set_high_mark(view_3D, element, KR_LIST, 0)[0]
        zero = self.set_high_mark(view_3D, element, KR_LIST, 0)[1]
        level = self.set_high_mark(view_3D, element, KR_LIST, 0)[2]

        element.LookupParameter('ФОП_Отметка от плиты КР').Set(kr_proximity)
        element.LookupParameter('ФОП_Отметка от нуля').Set(zero)
        element.LookupParameter('ФОП_Отметка от уровня').Set(level)

        width = element.LookupParameter('ФОП_РАЗМ_Ширина проёма').AsValueString()
        height = element.LookupParameter('ФОП_РАЗМ_Высота проёма').AsValueString()

        host = element.Host.get_Parameter(BuiltInParameter.ELEM_TYPE_PARAM).AsValueString()

        if "ГСП" in host:
            key = '[ГСП]' + '_[' + unique + ']_' + width + '_' + height + '_' + kr_proximity
            element.LookupParameter('ФОП_Группирование').Set(key)

        elif not any(IGNORE_WALL in host for IGNORE_WALL in IGNORE_WALLS):
            key = '[Кладка]' + '_[' + unique + ']_' + width + '_' + height + '_' + kr_proximity
            element.LookupParameter('ФОП_Группирование').Set(key)

        else:
            key = '[ЖБ]' + '_[' + unique + ']_' + width + '_' + height + '_' + kr_proximity
            element.LookupParameter('ФОП_Группирование').Set(key)
            element.LookupParameter('ФОП_Марка проема').Set('')

        return key

    # Получаем расстояние от вложенного семейства Ант_Маркер_Низ или Ант_Маркер_Верх до конструктивного элемента
    def set_high_mark(self, view_3D, element, list_of_kr, condition):
        # Получаем все вложенные семейства
        ids_subcomponents = element.GetSubComponentIds()

        for id in ids_subcomponents:
            subcomponent = doc.GetElement(id)
            # Ищем нужное семейство
            if condition == 0:
                if "Ант_Маркер_Низ" in subcomponent.Name:
                    marker = subcomponent
                    break
            else:
                if "Ант_Маркер_Верх" in subcomponent.Name:
                    marker = subcomponent
                    break

        # Получаем стартовые координаты луча
        start_point_x = marker.Location.Point.X
        start_point_y = marker.Location.Point.Y
        # Координату луча начинаем выше на 1 фут, чтобы учесть соприкасающиеся плоскости
        start_point_z = marker.Location.Point.Z + 0.1
        start_point_xyz = XYZ(start_point_x, start_point_y, start_point_z)

        # Получаем направление-координаты луча
        if condition == 0:
            start_point_z_dir = start_point_z - 100000000000
        else:
            start_point_z_dir = start_point_z + 100000000000

        direction_xyz = XYZ(start_point_x, start_point_y, start_point_z_dir)

        # создаем коллекцию ID
        collection_id = List[ElementId]([element.Host.Id])
        # создаем фильтр исключающий стену, в которой размещено окно, дверь или отверстие
        exclusion_filter = ExclusionFilter(collection_id)

        # создаем коллекцию фильтров по категориям - балки, перекрытия, стены, фундаменты
        collection_filter = List[BuiltInCategory](
            [BuiltInCategory.OST_StructuralFraming, BuiltInCategory.OST_StructuralFoundation,
             BuiltInCategory.OST_Floors, BuiltInCategory.OST_Walls])
        # создаем мультифильтр по категориям
        multicategory_filter = ElementMulticategoryFilter(collection_filter)
        # создаем фильтр по имени параметра
        # создаем параметр провайдер
        pvp = ParameterValueProvider((ElementId(BuiltInParameter.ALL_MODEL_TYPE_NAME)))
        # создаем условие
        fsre = FilterStringContains()
        # создаем правила для фильтра
        list_of_filters = []

        for kr_name in list_of_kr:
            rule = FilterStringRule(pvp, fsre, kr_name, True)
            parameter_filter = ElementParameterFilter(rule)
            list_of_filters.append(parameter_filter)

        # создаем логический фильтр куда передаем фильтры по правилам, чтобы перебрать все условия "или"
        logical_or_filter = LogicalOrFilter(list_of_filters)
        # создаем логический фильтр куда передаем мультикатегори-фильтр и фильтр по параметру
        logical_and_filter = LogicalAndFilter([multicategory_filter, logical_or_filter, exclusion_filter])
        # Создаем объект ReferenceIntersector
        ref_intersec = ReferenceIntersector(logical_and_filter, FindReferenceTarget.Face, view_3D)

        try:
            # Получаем первый элемент, который пересекает луч
            kr_element = ref_intersec.FindNearest(start_point_xyz, direction_xyz)

            # Дистанция до КР-элемента в миллиметрах (округленная)
            kr_distance = round((kr_element.Proximity - 0.1) * 304.8)

        except:

            kr_distance = 666

        location_zero = round((marker.Location.Point.Z) * 304.8) - (BASE_POINT[0].Position.Z * 304.8)

        location_level = self.convert_to_millimeters(marker.Location.Point.Z) - self.convert_to_millimeters(
            doc.GetElement(marker.LevelId).Elevation)

        return self.change_values(kr_distance), self.change_values(location_zero), self.change_values(location_level)


    def convert_to_millimeters(self, value):
        return UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters)


    # Меняем целое число на отметку формата +0.000
    def change_values(self, value):
        if "Обновите" in str(value):
            return value
        else:
            result_string = str(int(float(value)))

        if "-" in result_string:
            char = "-"
            replaced_string = result_string.replace('-', '')
        else:
            char = "+"
            replaced_string = result_string

        def add_nulls(string):
            if len(replaced_string) == 1:
                return "0,00"
            elif len(replaced_string) == 2:
                return "0,0"
            elif len(replaced_string) == 3:
                return "0,"
            else:
                return ","

        end_str = replaced_string[-3:]
        start_str = replaced_string[:(len(replaced_string) - len(end_str))]

        return char + start_str + add_nulls(replaced_string) + end_str


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
        self.__unique_identifier = "Кровля"
        self.__error_text = ""

        self.__create_command = CreateCommand(self)

    @reactive
    def unique_identifier(self):
        return self.__unique_identifier

    @unique_identifier.setter
    def unique_identifier(self, value):
        self.__unique_identifier = value

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
