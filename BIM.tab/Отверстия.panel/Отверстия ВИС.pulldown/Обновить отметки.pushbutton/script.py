#! /usr/bin/env python
# -*- coding: utf-8 -*-

import os.path as op
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
clr.AddReference("RevitAPI")
clr.AddReference("RevitAPIUI")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
import sys
import paraSpec
from Autodesk.Revit.DB import *
from Redomine import *
from System import Guid
from itertools import groupby
from pyrevit import revit, forms
from pyrevit.script import output

# colFittings = make_col(BuiltInCategory.OST_DuctFitting)
# colPipeFittings = make_col(BuiltInCategory.OST_PipeFitting)
# colPipeCurves = make_col(BuiltInCategory.OST_PipeCurves)
# colCurves = make_col(BuiltInCategory.OST_DuctCurves)
# colFlexCurves = make_col(BuiltInCategory.OST_FlexDuctCurves)
# colFlexPipeCurves = make_col(BuiltInCategory.OST_FlexPipeCurves)
# colTerminals = make_col(BuiltInCategory.OST_DuctTerminal)
# colAccessory = make_col(BuiltInCategory.OST_DuctAccessory)
# colPipeAccessory = make_col(BuiltInCategory.OST_PipeAccessory)
# colEquipment = make_col(BuiltInCategory.OST_MechanicalEquipment)
# colInsulations = make_col(BuiltInCategory.OST_DuctInsulations)
# colPipeInsulations = make_col(BuiltInCategory.OST_PipeInsulations)
# colPlumbingFixtures = make_col(BuiltInCategory.OST_PlumbingFixtures)
# colSprinklers = make_col(BuiltInCategory.OST_Sprinklers)
#
# collections = [colCurves, colPipeCurves, colSprinklers, colAccessory,
#                colPipeAccessory, colTerminals, colEquipment, colPlumbingFixtures]
#
# doc = __revit__.ActiveUIDocument.Document  # type: Document
# view = doc.ActiveView
#
#
# levelCol = make_col(BuiltInCategory.OST_Levels)
#
# class level:
#     def __init__(self, element):
#         self.name = element.Name
#         self.z = element.get_Parameter(BuiltInParameter.LEVEL_ELEV).AsValueString()
#
# report_rows = set()
#
# def getElementLevelName(element):
#     level = element.get_Parameter(BuiltInParameter.FAMILY_LEVEL_PARAM)
#     baseLevel = element.get_Parameter(BuiltInParameter.RBS_START_LEVEL_PARAM)
#
#     if element.Category.IsId(BuiltInCategory.OST_PipeInsulations) \
#             or element.Category.IsId(BuiltInCategory.OST_DuctInsulations):
#         return None
#     else:
#         if level:
#             level = level.AsValueString()
#         if baseLevel:
#             level = baseLevel.AsValueString()
#
#     return level
#
# def getElementBot(element):
#     if element.Category.IsId(BuiltInCategory.OST_DuctCurves):
#         bot = element.get_Parameter(BuiltInParameter.RBS_DUCT_BOTTOM_ELEVATION).AsValueString()
#         return bot
#
#     if element.Category.IsId(BuiltInCategory.OST_PipeCurves):
#         bot = element.get_Parameter(BuiltInParameter.RBS_PIPE_BOTTOM_ELEVATION).AsValueString()
#         return bot
#
#     bot = element.get_Parameter(BuiltInParameter.FLOOR_HEIGHTABOVELEVEL_PARAM).AsValueString()
#     return bot
#
#
#
# def getElementMid(element):
#     if element.Category.IsId(BuiltInCategory.OST_PipeCurves) or element.Category.IsId(BuiltInCategory.OST_DuctCurves):
#         mid = element.get_Parameter(BuiltInParameter.RBS_OFFSET_PARAM).AsValueString()
#         return mid
#     else:
#         return None
#
# def summ(x, y):
#     if ',' in x:
#         x = x.replace(',', '.')
#     if ',' in y:
#         y = y.replace(',', '.')
#     return float(x) + float(y)
#
# def execute():
#     with revit.Transaction("Обновление абсолютной отметки"):
#         projectLevels = []
#         for element in levelCol:
#             newLevel = level(element)
#             projectLevels.append(newLevel)
#
#         for collection in collections:
#             for element in collection:
#                 if not isElementEditedBy(element):
#                     midParam = element.LookupParameter('ФОП_ВИС_Отметка оси от нуля')
#                     botParam = element.LookupParameter('ФОП_ВИС_Отметка низа от нуля')
#                     elementLV = getElementLevelName(element)
#                     elementMid = getElementMid(element)
#                     elementBot = getElementBot(element)
#
#
#
#
#
#                     for projectLevel in projectLevels:
#                         if projectLevel.name == elementLV:
#
#                             if midParam:
#                                 midParam.Set(0)
#                             if midParam and elementMid:
#                                 markMid = summ(projectLevel.z, elementMid)
#                                 midParam.Set(markMid/1000)
#
#                             if botParam:
#                                 botParam.Set(0)
#                             if botParam and elementBot:
#                                 markBot = summ(projectLevel.z, elementBot)
#                                 botParam.Set(markBot/1000)
#
#                     try:
#                         if element.Host:
#                             elementBot = fromRevitToMilimeters(element.Location.Point[2])
#                             botParam.Set(elementBot/1000)
#                     except:
#                         pass
#                 else:
#                     fillReportRows(element,report_rows)
#
#     for report in report_rows:
#         print 'Некоторые элементы не были отработаны так как заняты пользователем ' + report
#
#
# if isItFamily():
#     print 'Надстройка не предназначена для работы с семействами'
#     sys.exit()
#
# parametersAdded = paraSpec.check_parameters()
#
# if not parametersAdded:
#     execute()

def get_elements_by_category(category):
    """ Возвращает коллекцию элементов по категории """
    col = FilteredElementCollector(doc)\
                            .OfCategory(category)\
                            .WhereElementIsNotElementType()\
                            .ToElements()
    return col

def convert(value):
    """ Преобразует дабл в миллиметры """
    new_v = UnitUtils.ConvertFromInternalUnits(value, UnitTypeId.Millimeters)
    return new_v


def get_parameter_if_exist_not_ro(element, built_in_parameters):
    """ Получает параметр, если он существует и если он не ReadOnly """
    for built_in_parameter in built_in_parameters:
        parameter = element.get_Parameter(built_in_parameter)
        if parameter is not None and not parameter.IsReadOnly:
            return built_in_parameter

    return None

def set_offset_value_pipe_elev(elements):
    for element in elements:
        # element.IsExistsParam(param) - True если существует
        # element.GetParamValue(param) - значение, если параметра нет - падаем
        # RBS_START_LEVEL_PARAM - базовый уровень
        element_level_id = element.GetParamValue(BuiltInParameter.RBS_START_LEVEL_PARAM)
        element_level = doc.GetElement(element_level_id)
        level_elev = element_level.GetParamValue(BuiltInParameter.LEVEL_ELEV)

        elem_mid_elevation = element.GetParamValue(BuiltInParameter.RBS_OFFSET_PARAM)

        elem_bot_elevation = element.GetParamValue(BuiltInParameter.RBS_PIPE_BOTTOM_ELEVATION)

        elem_mid_ablos_elev = level_elev + elem_mid_elevation
        elem_bot_ablos_elev = level_elev + elem_bot_elevation

        element.SetParamValue('ADSK_Отметка оси от нуля', elem_mid_ablos_elev)
        element.SetParamValue('ADSK_Отметка низа от нуля', elem_bot_ablos_elev)

def set_offset_value_duct_elev(elements):
    for element in elements:
        element_level_id = element.GetParamValue(BuiltInParameter.RBS_START_LEVEL_PARAM)
        element_level = doc.GetElement(element_level_id)
        level_elev = element_level.GetParamValue(BuiltInParameter.LEVEL_ELEV)

        elem_mid_elevation = element.GetParamValue(BuiltInParameter.RBS_OFFSET_PARAM)

        elem_bot_elevation = element.GetParamValue(BuiltInParameter.RBS_DUCT_BOTTOM_ELEVATION)

        elem_mid_ablos_elev = level_elev + elem_mid_elevation
        elem_bot_ablos_elev = level_elev + elem_bot_elevation

        element.SetParamValue('ADSK_Отметка оси от нуля', elem_mid_ablos_elev)
        element.SetParamValue('ADSK_Отметка низа от нуля', elem_bot_ablos_elev)

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

# Функция для поиска семейства в проекте
def find_family_symbol(family_name):
    family_symbols = FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()
    family_symbol = next((fs for fs in family_symbols if fs.Family.Name == family_name), None)

    return family_symbol

def check_family():
    family_names = ["ОбщМд_Отв_Отверстие_Прямоугольное_В стене",
                    "ОбщМд_Отв_Отверстие_Прямоугольное_В перекрытии",
                    "ОбщМд_Отв_Отверстие_Круглое_В перекрытии",
                    "ОбщМд_Отв_Отверстие_Круглое_В стене"]

    param_list = [
        "ADSK_Отверстие_ОтметкаЭтажа",
        "ADSK_Отверстие_ОтметкаОтНуля",
        "ADSK_Отверстие_ОтметкаОтЭтажа"
    ]
    for family_name in family_names:
        family = find_family_symbol(family_name).Family

        if family is None:
            return

        symbol_params = get_family_shared_parameter_names(family)
        for param in param_list:
            if param not in symbol_params:
                forms.alert("Параметра {} нет в семействах отверстий. Обновите все семейства отверстий из базы семейств.".
                            format(param), "Ошибка", exitscript=True)

def set_offset_value_generic_elev(elements):
    for element in elements:
        symbol = element.Symbol
        family_name = symbol.GetParamValue(BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM)

        if family_name in  ["ОбщМд_Отв_Отверстие_Прямоугольное_В стене",
                    "ОбщМд_Отв_Отверстие_Прямоугольное_В перекрытии",
                    "ОбщМд_Отв_Отверстие_Круглое_В перекрытии",
                    "ОбщМд_Отв_Отверстие_Круглое_В стене"]:
            element_level_id = element.GetParamValue(BuiltInParameter.FAMILY_LEVEL_PARAM)
            element_level = doc.GetElement(element_level_id)
            level_elev = round(convert(element_level.GetParamValue(BuiltInParameter.LEVEL_ELEV)), 2)
            element.SetParamValue('ADSK_Отверстие_ОтметкаЭтажа', level_elev)

            elem_current_elevation = round(convert(element.GetParamValue(BuiltInParameter.INSTANCE_ELEVATION_PARAM)), 2)
            element.SetParamValue('ADSK_Отверстие_ОтметкаОтЭтажа', elem_current_elevation)

            elem_ablos_elevation = level_elev + elem_current_elevation
            element.SetParamValue('ADSK_Отверстие_ОтметкаОтНуля', elem_ablos_elevation)



def execute():
    duct_col = get_elements_by_category(BuiltInCategory.OST_DuctCurves)
    pipe_col = get_elements_by_category(BuiltInCategory.OST_PipeCurves)
    generic_col = get_elements_by_category(BuiltInCategory.OST_GenericModel)
    check_family()
    with revit.Transaction("BIM: Обновление отметки"):
        set_offset_value_pipe_elev(pipe_col)
        set_offset_value_duct_elev(duct_col)
        set_offset_value_generic_elev(generic_col)

execute()