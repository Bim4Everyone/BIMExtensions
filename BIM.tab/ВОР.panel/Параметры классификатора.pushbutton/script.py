# -*- coding: utf-8 -*-
import os
import clr
import datetime
from System.Collections.Generic import *

clr.AddReference("Microsoft.Office.Interop.Excel")
import Microsoft.Office.Interop.Excel as Excel
from System.Runtime.InteropServices import Marshal

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")
import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)
from dosymep_libs.bim4everyone import *

import pyevent
from pyrevit import EXEC_PARAMS, revit
from pyrevit.forms import *
from pyrevit import script

import Autodesk.Revit.DB
from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

chapter_parameter = "ФОП_МТР_Наименование главы"
work_title_parameter = "ФОП_МТР_Наименование работы"
unit_parameter = "ФОП_МТР_Единица измерения"
calculation_type_parameter = "ФОП_МТР_Тип подсчета"

calculation_type_dict = {
    "м": 1,
    "м²": 2,
    "м³": 3,
    "шт.": 4}

default_excel_path = "W:\Проектный институт\Проектные Группы\Типовые ТЗ\BIM-стандарт A101\Классификатор видов работ.xlsx"

report_no_work_code = []
report_classifier_code_not_found = []
report_edited = []
report_not_edited = []
report_errors = []


# Класс для хранения инфы по видам работ из Excel
class Work:
    def __init__(self, code, chapter, title_of_work, unit_of_measurement):
        self.code = code  # код работы
        self.chapter = chapter  # наименование главы
        self.title_of_work = title_of_work  # наименование работы
        self.unit_of_measurement = unit_of_measurement  # единица измерения


# Класс-оболочка для хранения информации
class RevitMaterial:
    def __init__(self, keynote, material, work):
        self.keynote = keynote  # ключевая заметка материала из Revit
        self.material = material  # материал из Revit
        self.work = work  # работа, параметры которой нужно назначить


def read_from_excel(path):
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
        dict_for_data = {}
        code = ""
        chapter = ""
        title_of_work = ""
        unit_of_measurement = ""
        for i in range(2, row_end_1 + 1):
            unit_of_measurement = ws_1.Cells(i, 3).Text

            if not unit_of_measurement:
                chapter = ws_1.Cells(i, 2).Text
                continue
            else:
                code = ws_1.Cells(i, 1).Text
                title_of_work = ws_1.Cells(i, 2).Text
                work = Work(code, chapter, title_of_work, unit_of_measurement)
                dict_for_data[code] = work
    except:
        output = script.output.get_output()
        output.close()
        alert("При чтении Excel-файла Классификатора произошла ошибка", exitscript=True)
    finally:
        excel.ActiveWorkbook.Close(False)
        Marshal.ReleaseComObject(ws_1)
        Marshal.ReleaseComObject(workbook)
        Marshal.ReleaseComObject(excel)
    return dict_for_data


def get_calculation_type_value(unit_value):
    if calculation_type_dict.has_key(unit_value):
        return calculation_type_dict[unit_value]
    else:
        return "Ошибка"


def set_param(param, value, edited):
    if param.AsValueString() == str(value):
        return edited
    else:
        param.Set(value)
        edited = True
        return edited


def set_classifier_parameters(revit_materials):
    with revit.Transaction("BIM: Заполнение параметров классификатора"):
        try:
            for revit_material in revit_materials:
                edited = False
                material = revit_material.material
                work = revit_material.work

                edited = set_param(
                    material.GetParam(chapter_parameter),
                    work.chapter,
                    edited)

                edited = set_param(
                    material.GetParam(work_title_parameter),
                    work.title_of_work,
                    edited)

                edited = set_param(
                    material.GetParam(unit_parameter),
                    work.unit_of_measurement,
                    edited)

                calculation_type = get_calculation_type_value(work.unit_of_measurement)
                edited = set_param(
                    material.GetParam(calculation_type_parameter),
                    calculation_type,
                    edited)

                if edited:
                    report_edited.append(["ИЗМЕНЕН", revit_material.keynote, material.Name])
                else:
                    report_not_edited.append(["БЕЗ ИЗМЕНЕНИЙ", revit_material.keynote, material.Name])
        except:
            report_errors.append(["ОШИБКА ПРИ ЗАПИСИ", revit_material.keynote, material.Name])


def get_excel_path():
    excel_path = default_excel_path
    if not os.path.exists(excel_path):
        excel_path = pick_file(
            files_filter="excel files (*.xlsx)|*.xlsx",
            init_dir="c:\\",
            restore_dir=True,
            multi_file=False,
            unc_paths=False,
            title="Выберите excel-файл Классификатора")

    if not excel_path:
        output = script.output.get_output()
        output.close()
        alert("Не указан путь к Excel-файлу Классификатора", exitscript=True)
    return excel_path


def get_materials():
    elems = (FilteredElementCollector(doc, doc.ActiveView.Id)
             .WhereElementIsNotElementType()
             .ToElements())

    if len(elems) == 0:
        output = script.output.get_output()
        output.close()
        alert("На активном виде не найдено ни одного элемента", exitscript=True)

    materialIds = []
    for elem in elems:
        for materialId in elem.GetMaterialIds(False):
            if materialId not in materialIds:
                materialIds.append(materialId)

    materials = []
    for materialId in materialIds:
        material = doc.GetElement(materialId)
        materials.append(material)

    if len(materials) == 0:
        output = script.output.get_output()
        output.close()
        alert("На активном виде не найдено ни одного элемента, " +
              "у которого можно забрать материал", exitscript=True)
    return materials

def get_report():
    report_part_1 = sorted(report_errors, key=lambda report_item: report_item[1])
    report_part_2 = sorted(report_edited, key=lambda report_item: report_item[1])
    report_part_3 = sorted(report_not_edited, key=lambda report_item: report_item[1])
    report_part_4 = sorted(report_no_work_code, key=lambda report_item: report_item[1])
    report_part_5 = sorted(report_classifier_code_not_found, key=lambda report_item: report_item[1])

    report = (report_part_1
              + report_part_2
              + report_part_3
              + report_part_4
              + report_part_5)
    return report


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Здравствуйте! Данный плагин предназначен для записи в параметры материалов "
          + "информации из Классификатора видов работ.")

    print("Собираю материалы у элементы на активном виде...")
    materials = get_materials()
    print("Найдено материалов: " + str(len(materials)))

    excel_path = get_excel_path()
    print("Читаю excel-файл Классификатора видов работ по пути: " + excel_path)

    dict_from_excel = read_from_excel(excel_path)
    print("Найдено видов работ: " + str(len(dict_from_excel.keys())))

    revit_materials = []
    for material in materials:
        keynote = material.GetParamValueOrDefault(BuiltInParameter.KEYNOTE_PARAM)

        # Отсеиваем ситуации, когда у материала не указана Ключевая заметка (код работы)
        if not keynote:
            report_no_work_code.append(["НЕТ КОДА РАБОТЫ", "", material.Name])
            continue

        # Отсеиваем ситуации, когда Классификатор не содержит указанный в материале код
        if not dict_from_excel.has_key(keynote):
            report_classifier_code_not_found.append(["НЕ НАЙДЕН КОД", keynote, material.Name])
            continue

        revit_materials.append(RevitMaterial(keynote, material, dict_from_excel[keynote]))

    set_classifier_parameters(revit_materials)

    report = get_report()
    output = script.output.get_output()
    output.print_table(table_data=report,
                       title="Отчет работы плагина",
                       columns=["Статус⠀⠀⠀⠀⠀⠀⠀⠀", "Код работы", "Имя материала"],
                       formats=['', '', ''])


script_execute()
