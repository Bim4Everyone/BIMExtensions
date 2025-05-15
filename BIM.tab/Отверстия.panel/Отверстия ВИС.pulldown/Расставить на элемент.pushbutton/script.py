#! /usr/bin/env python
# -*- coding: utf-8 -*-


import clr
import datetime
import os
import json
import codecs
from pyrevit.labs import target
from pyrevit.revit.db.query import get_family_symbol
from unicodedata import category

clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import sys
import System
import math
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.UI.Selection import ISelectionFilter
from Autodesk.Revit.DB import (BuiltInCategory,
                               ElementFilter,
                               LogicalOrFilter,
                               ElementCategoryFilter,
                               ElementMulticategoryFilter,
                               Line,
                               ElementTransformUtils)

from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List
from System import Guid
from System import Environment
from pyrevit import forms
from pyrevit import revit
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS


from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document # type: Document

uidoc = __revit__.ActiveUIDocument
uiapp = __revit__.Application

view = doc.ActiveView

# Создаем пользовательский фильтр для выбора объектов
class CustomSelectionFilter(ISelectionFilter):
    def __init__(self, filter):
        self.filter = filter

    def AllowElement(self, element):
        return self.filter.PassesFilter(element) and not self.IsOval(element)

    def AllowReference(self, reference, position):
        return True

    def IsOval(self, element):
        if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
            return element.DuctType.Shape == ConnectorProfileType.Oval

    def IsVertical(self, element):
        start_xyz, end_xyz = get_connector_coordinates(element)

        # Вычисляем разности координат
        delta_x = round(end_xyz.X, 3) - round(start_xyz.X, 3)
        delta_y = round(end_xyz.Y, 3) - round(start_xyz.Y, 3)
        delta_z = round(end_xyz.Z, 3) - round(start_xyz.Z, 3)
        epsilon = 0.01 # Соответствует смещению в 3мм

        # Если линия вертикальна (delta_x < epsilon и delta_y < epsilon), возвращаем True
        if abs(delta_x) < epsilon and abs(delta_y) < epsilon:
            return True

        if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
            slope = element.GetParamValue(BuiltInParameter.RBS_PIPE_SLOPE)
            if slope > 0.06:
                return True
        else:
            if abs(delta_z) > epsilon:
                return True

        return False

# Определяем класс для хранения данных из конфига отверстий
class CategoryConfig:
    def __init__(self, category_name, from_value, to_value, offset_value, opening_type_name, step):
        self.category_name = category_name
        self.from_value = from_value
        self.to_value = to_value
        self.offset_value = offset_value
        self.opening_type_name = opening_type_name
        self.step = step

# Класс для передачи в качестве задачи на создание экземпляра
class Objective:
    family_name = None
    indent = None
    curve = None
    curve_width = None
    curve_height = None
    curve_level = None
    category_name = None
    point = None
    direction = None
    angle = 0
    family_symbol = None
    is_vertical = None

    instance = None
    step = None

    def is_curve_vertical(self, element):
        start_xyz, end_xyz = get_connector_coordinates(element)

        # Вычисляем разности координат
        delta_x = round(end_xyz.X, 3) - round(start_xyz.X, 3)
        delta_y = round(end_xyz.Y, 3) - round(start_xyz.Y, 3)
        delta_z = round(end_xyz.Z, 3) - round(start_xyz.Z, 3)
        epsilon = 0.01 # Соответствует смещению в 3мм

        # Если линия вертикальна (delta_x < epsilon и delta_y < epsilon), возвращаем True
        if abs(delta_x) < epsilon and abs(delta_y) < epsilon:
            return True

        if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
            slope = element.GetParamValue(BuiltInParameter.RBS_PIPE_SLOPE)
            if slope > 0.06:
                return True
        else:
            if abs(delta_z) > epsilon:
                return True

        return False

    def __init__(self, curve, point):
        self.curve = curve
        self.point = point
        self.curve_level = get_curve_level(self.curve)
        self.curve_width, self.curve_height, self.category_name = get_curve_characteristic(self.curve)
        self.direction = get_curve_direction(self.curve)
        self.is_vertical = self.is_curve_vertical(self.curve)
        if self.is_vertical:
            self.angle = get_curve_XY_ange(self.curve)
        self.indent, self.family_name, self.step = get_plugin_config(self.curve, self.is_vertical)
        self.family_symbol = find_family_symbol(self.family_name)

# Получаем место клика
def get_click_reference():
    categories = List[BuiltInCategory]()

    categories.Add(BuiltInCategory.OST_DuctCurves)
    categories.Add(BuiltInCategory.OST_PipeCurves)
    categories.Add(BuiltInCategory.OST_CableTray)
    categories.Add(BuiltInCategory.OST_Conduit)

    # Создаем MultiCategoryFilter для этих категорий
    multi_category_filter = ElementMulticategoryFilter(categories)

    # Применяем пользовательский фильтр к выбору объектов
    selection_filter = CustomSelectionFilter(multi_category_filter)

    try:
        reference = uidoc.Selection.PickObject(ObjectType.Element, selection_filter)
    except Autodesk.Revit.Exceptions.OperationCanceledException:
        sys.exit()

    return reference

# Функция для получения центра и направления воздуховода
def get_curve_direction(duct):
    options = Options()
    geometry = duct.get_Geometry(options)
    bbox = geometry.GetBoundingBox()
    # Получение направления воздуховода
    curve = next((geom for geom in geometry if isinstance(geom, Solid) and geom.Faces.Size > 0), None)
    if curve:
        face = curve.Faces.get_Item(0)
        normal = face.ComputeNormal(UV(0.5, 0.5))
        direction = normal.CrossProduct(XYZ.BasisZ)
    else:
        direction = XYZ.BasisX  # По умолчанию используем направление по оси X
    return direction

# Получает координатную систему для линейного элемента
def get_connector_coordinate_system(mep_curve):
    connector_manager = mep_curve.ConnectorManager
    for connector in connector_manager.Connectors:
        if connector.CoordinateSystem:
            return connector.CoordinateSystem
    return None

# Получает угол XY-для поворта отверстий на угол поворта линейного элемента
def get_curve_XY_ange(curve):
    transform = get_connector_coordinate_system(curve)
    angle_rad = XYZ(1, 0, 0).AngleOnPlaneTo(transform.BasisX, XYZ(0, 0, 1))  # плоскость XY = нормаль Z

    return angle_rad

# Функция для поиска семейства в проекте
def find_family_symbol(family_name):
    family_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    family_symbol = next((fs for fs in family_symbols if fs.Family.Name == family_name), None)
    if not family_symbol:
        forms.alert("Семейства отверстий не найдены", "Ошибка", exitscript=True)

    return family_symbol

def check_family():
    family_names = [rectangle_opening_floor_name,
                    rectangle_opening_wall_name,
                    round_opening_floor_name,
                    round_opening_wall_name]

    param_list = [
        shared_currency_level_offset_name,
        shared_currency_from_level_offset_name,
        shared_currency_absolute_offset_name,
        shared_level_offset_name,
        shared_from_level_offset_name,
        shared_absolute_offset_name
        ]

    for family_name in family_names:
        family = find_family_symbol(family_name).Family
        symbol_params = get_family_shared_parameter_names(family)

        for param in param_list:
            if param not in symbol_params:
                forms.alert("Параметра {} нет в семействах отверстий. Обновите все семейства отверстий из базы семейств.".
                            format(param), "Ошибка", exitscript=True)

# Функция для получения размеров элемента
def get_element_size(element):
    options = Options()
    geometry = element.get_Geometry(options)
    bbox = geometry.GetBoundingBox()
    size = bbox.Max - bbox.Min
    return size

def get_connector_coordinates(element):
    # Получаем коннекторы воздуховода
    connectors = element.ConnectorManager.Connectors

    # Получаем координаты начала и конца воздуховода через коннекторы
    start_point = None
    end_point = None

    for connector in connectors:
        # К линейному элементу может быть подключено произвольное количество врезок, каждая из которых попадет
        # в список коннекторов линейного элемента. Проверяем чтоб айди владельца был айди рабочего элемента
        if connector.Owner.Id == element.Id:
            if start_point is None:
                start_point = connector.Origin
            else:
                end_point = connector.Origin
                break

    if start_point is None or end_point is None:
        forms.alert("Не удалось получить координаты коннекторов.", "Ошибка", exitscript=True)

    # Получаем координаты начала и конца воздуховода
    start_xyz = XYZ(start_point.X, start_point.Y, start_point.Z)
    end_xyz = XYZ(end_point.X, end_point.Y, end_point.Z)

    return start_xyz, end_xyz

# Получаем точку контакта нормали от точки клика и оси линейного элемента
def get_contact_point(element, point):
    # Получаем координаты начальной и конечной точек линейного элемента
    start_xyz, end_xyz = get_connector_coordinates(element)

    # Вектор направления прямой
    line_vector = end_xyz - start_xyz

    # Вектор от начальной точки прямой до точки
    point_vector = point - start_xyz

    # Проекция вектора на вектор направления прямой
    projection = point_vector.DotProduct(line_vector) / line_vector.DotProduct(line_vector)

    # Координаты точки соприкосновения
    contact_point = start_xyz + line_vector * projection

    return contact_point

# Возвращает параметр или None
def get_parameter_if_exists(element, param_name):
    if element.IsExistsParam(param_name):
        return element.GetParam(param_name)
    else:
        return None

# Округляет до ближайшего указанного числа
def round_up_to_nearest(number, step):
    step = UnitUtils.ConvertToInternalUnits(step, UnitTypeId.Millimeters)
    remainder = number % step

    if remainder == 0:
        return number
    else:
        return number + (step - remainder)

# Устанавливаем размер круглого отверстия
def set_size_round_opening(instance_diameter_param, objective):
    instance_diameter_param.Set(round_up_to_nearest(
        max(objective.curve_width, objective.curve_height) + objective.indent, objective.step))
    # для круглых отверстий не нужна поправка, у них точка вставки в центре
    return 0

# Устанавливаем размер прямоугольного отверстия
def set_size_rectangular_opening(instance_width_param, instance_height_param, objective):
    instance_width_param.Set(round_up_to_nearest(objective.curve_width + objective.indent, objective.step) )

    rounded_height = round_up_to_nearest(objective.curve_height + objective.indent, objective.step)
    instance_height_param.Set(rounded_height )

    # для прямоугольных точка вставки на основании, нужно их смещать на половину высоты
    if config_wall_type_name in objective.family_name:
        return rounded_height / 2
    else:
        return 0

# Получаем характеристику линейного элемента, его габарит и имя
def get_curve_characteristic(curve):
    def get_pipe_dimensions(curve):
        width = curve.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
        height = curve.GetParamValue(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER)
        return width, height, config_category_pipe_name

    def get_duct_dimensions(curve):
        if curve.DuctType.Shape == ConnectorProfileType.Round:
            width = curve.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            height = curve.GetParamValue(BuiltInParameter.RBS_CURVE_DIAMETER_PARAM)
            return width, height, config_category_round_duct_name
        elif curve.DuctType.Shape == ConnectorProfileType.Rectangular:
            width = curve.GetParamValue(BuiltInParameter.RBS_CURVE_WIDTH_PARAM)
            height = curve.GetParamValue(BuiltInParameter.RBS_CURVE_HEIGHT_PARAM)
            return width, height, config_category_rectangle_duct_name
        return 0, 0, None

    def get_cable_tray_dimensions(curve):
        width = curve.GetParamValue(BuiltInParameter.RBS_CABLETRAY_WIDTH_PARAM)
        height = curve.GetParamValue(BuiltInParameter.RBS_CABLETRAY_HEIGHT_PARAM)
        return width, height, config_category_trays_name

    def get_conduit_dimensions(curve):
        width = curve.GetParamValue(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
        height = curve.GetParamValue(BuiltInParameter.RBS_CONDUIT_OUTER_DIAM_PARAM)
        return width, height, config_category_conduit_name

    if curve.Category.IsId(BuiltInCategory.OST_PipeCurves):
        return get_pipe_dimensions(curve)
    elif curve.Category.IsId(BuiltInCategory.OST_DuctCurves):
        return get_duct_dimensions(curve)
    elif curve.Category.IsId(BuiltInCategory.OST_CableTray):
        return get_cable_tray_dimensions(curve)
    elif curve.Category.IsId(BuiltInCategory.OST_Conduit):
        return get_conduit_dimensions(curve)

    return 0, 0, None

# Настройка размера под размер линейного элемента. Возвращает корректировку по вертикали для точки вставки
def setup_size(objective):
    # Возвращает параметры, отвечающие за размерности экземпляра отверстия
    def get_instance_parameters(instance):
        instance_diameter_param = get_parameter_if_exists(instance, shared_diameter_param_name)
        instance_height_param = get_parameter_if_exists(instance, shared_height_param_name)
        instance_width_param = get_parameter_if_exists(instance, shared_width_param_name)
        return instance_diameter_param, instance_height_param, instance_width_param

    # Настройка размеров под короба
    def handle_pipe_or_conduit(instance_params, objective):
        instance_diameter_param, instance_height_param, instance_width_param = instance_params
        if config_round_type_name in objective.family_name:
            return set_size_round_opening(instance_diameter_param,
                                          objective)
        if config_rectangle_type_name in objective.family_name:
            return set_size_rectangular_opening(instance_width_param,
                                                instance_height_param,
                                                objective)

    # Настройка размера отверстия под линейный элемент
    def handle_duct_curves(instance_params, objective):
        instance_diameter_param, instance_height_param, instance_width_param = instance_params
        if objective.curve.DuctType.Shape == ConnectorProfileType.Round:
            if config_round_type_name in objective.family_name:
                return set_size_round_opening(
                    instance_diameter_param,
                    objective)

            if config_rectangle_type_name in objective.family_name:
                return set_size_rectangular_opening(
                    instance_width_param,
                    instance_height_param,
                    objective)

        if objective.curve.DuctType.Shape == ConnectorProfileType.Rectangular:
            if config_round_type_name in objective.family_name:
                return set_size_round_opening(
                    instance_diameter_param,
                    math.sqrt(objective.curve_width ** 2 + objective.curve_height ** 2))

            if config_rectangle_type_name in objective.family_name:
                return set_size_rectangular_opening(
                    instance_width_param,
                    instance_height_param,
                    objective)

    # Настройка размера под кабельный лоток
    def handle_cable_tray(instance_params, objective):
        instance_diameter_param, instance_height_param, instance_width_param = instance_params
        if config_round_type_name in objective.family_name:
            return set_size_round_opening(
                instance_diameter_param,
                math.sqrt(objective.curve_width ** 2 + objective.curve_height ** 2))

        if config_rectangle_type_name in objective.family_name:
            return set_size_rectangular_opening(
                instance_width_param,
                instance_height_param,
                objective)

    instance_params = get_instance_parameters(objective.instance)

    if (objective.curve.Category.IsId(BuiltInCategory.OST_PipeCurves)
            or objective.curve.Category.IsId(BuiltInCategory.OST_Conduit)):
        return handle_pipe_or_conduit(instance_params, objective)

    if objective.curve.Category.IsId(BuiltInCategory.OST_DuctCurves):
        return handle_duct_curves(instance_params, objective)

    if objective.curve.Category.IsId(BuiltInCategory.OST_CableTray):
        return handle_cable_tray(instance_params, objective)

# Возвращаем значение системного Имя системы или ФОП_ВИС_Имя системы в зависимости от заполненности второго
def get_curve_system(curve):
    system_name = curve.GetParamValueOrDefault(shared_system_param_name)
    if system_name is None:
        system_name = curve.GetParamValueOrDefault(BuiltInParameter.RBS_SYSTEM_NAME_PARAM)

    if system_name is None:
        return "Нет системы"
    return system_name

def set_offset_values_to_shared_params(instance, curve_level):
    absolute_offset = instance.GetParamValue(BuiltInParameter.INSTANCE_ELEVATION_PARAM) + curve_level.Elevation
    from_level_offset = instance.GetParamValue(BuiltInParameter.INSTANCE_ELEVATION_PARAM)
    level_offset = curve_level.Elevation

    instance.SetParamValue(shared_absolute_offset_name,
                            absolute_offset)
    instance.SetParamValue(shared_from_level_offset_name,
                           from_level_offset)
    instance.SetParamValue(shared_level_offset_name,
                           level_offset)

    # Currency - денежная единица. У нее нет размерности и нужно просто дабл в нее подать, поэтому преобразуем
    absolute_offset =  UnitUtils.ConvertFromInternalUnits(absolute_offset, UnitTypeId.Millimeters)
    from_level_offset = UnitUtils.ConvertFromInternalUnits(from_level_offset, UnitTypeId.Millimeters)
    level_offset= UnitUtils.ConvertFromInternalUnits(level_offset, UnitTypeId.Millimeters)

    instance.SetParamValue(shared_currency_absolute_offset_name,
                           absolute_offset)
    instance.SetParamValue(shared_currency_from_level_offset_name,
                           from_level_offset)
    instance.SetParamValue(shared_currency_level_offset_name,
                           level_offset)

#Возвращает уровень линейного элемента
def get_curve_level(curve):
    all_levels = FilteredElementCollector(doc).OfClass(Level).ToElements()
    curve_level_name = curve.GetParam(BuiltInParameter.RBS_START_LEVEL_PARAM).AsValueString()

    for level in all_levels:
        if curve_level_name == level.Name:
            return level

# Функция для чтения JSON файла и извлечения данных
def get_category_configs(file_path, category_names):
    with codecs.open(file_path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    category_configs = []
    for category in data.get('Categories', []):
        if category.get('Name') in category_names:

            for offset in category.get('Offsets', []):
                category_configs.append(CategoryConfig(
                    category_name=category.get('Name'),
                    from_value=offset.get('From'),
                    to_value=offset.get('To'),
                    offset_value=offset.get('OffsetValue'),
                    opening_type_name=offset.get("OpeningTypeName"),
                    step=category.get("Rounding")
                ))

    return category_configs

# Получаем конфиг плагина размещение отверстий или возвращаем стандартные настройки
def get_plugin_config(curve, is_vertical):
    indent = UnitUtils.ConvertToInternalUnits((50 * 2), UnitTypeId.Millimeters)
    family_name = rectangle_opening_wall_name
    step = 50

    curve_width, curve_height, category_name = get_curve_characteristic(curve)
    curve_size = UnitUtils.ConvertToInternalUnits(max(curve_width, curve_height), UnitTypeId.Millimeters)

    version = uiapp.VersionNumber
    documents_path = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

    path_settings_file_path = os.path.join(documents_path,
                                         'dosymep',
                                         str(version),
                                         'RevitOpeningPlacement',
                                         'MepConfigPath.json')
    if os.path.isfile(path_settings_file_path):
        with codecs.open(path_settings_file_path, 'r', encoding='utf-8') as data_file:
            data = json.load(data_file)

        configuration_file_path = data.get("OpeningConfigPath")

    else:
        configuration_file_path = os.path.join(documents_path,
                                             'dosymep',
                                             str(version),
                                             'RevitOpeningPlacement',
                                             'OpeningConfig.json')

    if os.path.isfile(configuration_file_path):
        category_names = [config_category_pipe_name,
                          config_category_round_duct_name,
                          config_category_rectangle_duct_name,
                          config_category_trays_name,
                          config_category_conduit_name]

        category_configs = get_category_configs(configuration_file_path, category_names)
        for config in category_configs:
            if category_name == config.category_name and config.from_value <= curve_size <= config.to_value:
                if config.opening_type_name == config_round_type_name:
                    if is_vertical:
                        family_name = round_opening_floor_name
                    else:
                        family_name = round_opening_wall_name
                if config.opening_type_name == config_rectangle_type_name:
                    if is_vertical:
                        family_name = rectangle_opening_floor_name
                    else:
                        family_name = rectangle_opening_wall_name

                step = config.step
                indent = UnitUtils.ConvertToInternalUnits(config.offset_value*2, UnitTypeId.Millimeters)

    return indent, family_name, step

def rotate_element(objective):
    if objective.angle == 0:
        return

    axis = Line.CreateUnbound(objective.point, XYZ(0, 0, 1))
    ElementTransformUtils.RotateElement(doc, objective.instance.Id, axis, objective.angle)

# Настраиваем экземпляр отверстия для чтения через навигатор АР и ставим его размеры-смещение под линейный элемент
def setup_opening_instance(objective):
    # Заполняем автора задания
    user_name = __revit__.Application.Username
    objective.instance.SetParamValue(shared_autor_param_name, user_name)

    # Заполняем айди линейного элемента
    objective.instance.SetParamValue(shared_info_param_name, objective.curve.Name +
                                     ": " + objective.curve.Id.ToString())

    # Заполняем имя системы элемента
    objective.instance.SetParamValue(shared_system_param_name, get_curve_system(objective.curve))

    # Заполняем время
    current_time = datetime.datetime.now()
    formatted_time = current_time.strftime("%Y-%m-%d %H:%M")
    objective.instance.SetParamValue(shared_date_param_name, formatted_time)

    # Устанавливаем размер под воздуховод и получаем смещение относительно него
    correction = setup_size(objective)

    rotate_element(objective)

    instance_offset_param = objective.instance.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)
    instance_current_offset = objective.instance.GetParamValueOrDefault(BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM)

    # перезаписываем отметку с учетом полученной корректировки и с округлением до ближайших 10
    if correction:
        instance_offset_param.Set(round_up_to_nearest(instance_current_offset - correction, 10))

    # Теоретически может быть и None, тогда просто размещаем
    if objective.curve_level is not None:
        # Заполняем параметры отметки от уровня и абсолютной отметки
        set_offset_values_to_shared_params(objective.instance, objective.curve_level)

# Функция для размещения семейства в заданных координатах и разворота вдоль оси линейного элемента
def place_family_at_coordinates(objective):
    # Сдвиг точки размещения на ось воздуховода по вертикали
    objective.point = get_contact_point(objective.curve, objective.point)

    # Создание экземпляра
    objective.family_symbol.Activate()

    if objective.curve_level is None:
        objective.instance = doc.Create.NewFamilyInstance(
            objective.point,
            objective.family_symbol,
            Structure.StructuralType.NonStructural)
    else:
        # Корректируем уровень по Z, потому что изначально точку мы получаем проектную и иначе она будет значительно выше чем должна
        objective.point = XYZ(
            objective.point.X,
            objective.point.Y,
            objective.point.Z - objective.curve_level.ProjectElevation)
        objective.instance = doc.Create.NewFamilyInstance(
            objective.point,
            objective.family_symbol,
            objective.curve_level,
            Structure.StructuralType.NonStructural)

    # Регенирируем документ, должно спасать от бага с присвоением не тех значений параметров
    doc.Regenerate()

    # Создание оси вращения, проходящей через точку размещения и направленной вдоль оси Z
    axis = Line.CreateBound(objective.point, objective.point + XYZ.BasisZ)

    # Вычисление угла поворота
    angle = math.atan2(objective.direction.Y, objective.direction.X)

    # Вращение экземпляра семейства вокруг оси Z
    objective.instance.Location.Rotate(axis, angle)

    return objective

def get_family_shared_parameter_names(family):
    # Открываем документ семейства для редактирования
    family_doc = doc.EditFamily(family)

    shared_parameters = []
    try:
        # Получаем менеджер семейства
        family_manager = family_doc.FamilyManager

        # Получаем все параметры семейства
        parameters = family_manager.GetParameters()

        # Фильтруем параметры, чтобы оставить только общие
        shared_parameters = [param.Definition.Name for param in parameters if param.IsShared]

        return shared_parameters
    finally:
        # Закрываем документ семейства без сохранения изменений
        family_doc.Close(False)

rectangle_opening_wall_name = "ОбщМд_Отв_Отверстие_Прямоугольное_В стене"
round_opening_wall_name = "ОбщМд_Отв_Отверстие_Круглое_В стене"

rectangle_opening_floor_name = "ОбщМд_Отв_Отверстие_Прямоугольное_В перекрытии"
round_opening_floor_name = "ОбщМд_Отв_Отверстие_Круглое_В перекрытии"

config_round_type_name = "Круглое"
config_rectangle_type_name = "Прямоугольное"
config_wall_type_name = "В стене"
config_floor_type_name = "В перекрытии"
config_category_pipe_name = "Трубы"
config_category_round_duct_name = "Воздуховоды (круглое сечение)"
config_category_rectangle_duct_name = "Воздуховоды (прямоугольное сечение)"
config_category_trays_name = "Лотки"
config_category_conduit_name = "Короба"
shared_currency_absolute_offset_name = "ADSK_Отверстие_ОтметкаОтНуля"
shared_currency_from_level_offset_name = "ADSK_Отверстие_ОтметкаОтЭтажа"
shared_currency_level_offset_name = "ADSK_Отверстие_ОтметкаЭтажа"
shared_absolute_offset_name = "ADSK_Отверстие_Отметка от нуля"
shared_from_level_offset_name = "ADSK_Отверстие_Отметка от этажа"
shared_level_offset_name = "ADSK_Отверстие_Отметка этажа"
shared_height_param_name = "ADSK_Размер_Высота"
shared_width_param_name = "ADSK_Размер_Ширина"
shared_diameter_param_name = "ADSK_Размер_Диаметр"
shared_date_param_name = "ФОП_Дата"
shared_info_param_name = "ФОП_Описание"
shared_autor_param_name = "ФОП_Автор задания"
shared_system_param_name = SharedParamsConfig.Instance.VISSystemName.Name


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    check_family()

    reference = get_click_reference()

    objective = Objective(doc.GetElement(reference), reference.GlobalPoint)

    with revit.Transaction("BIM: Добавление отверстия"):
        # Размещение и поворот
        objective = place_family_at_coordinates(objective)
        # Настройка размеров и параметров отверстия
        setup_opening_instance(objective)

script_execute()