# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep.Revit
clr.ImportExtensions(dosymep.Revit)

from itertools import groupby

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

from dosymep.Bim4Everyone.Templates import ProjectParameters

from pyrevit import script
from pyrevit import forms

output = script.get_output()
output.set_title("dosymep (Обновление номера вида)")
output.center()

application = __revit__.Application
document = __revit__.ActiveUIDocument.Document


class Section(object):
    def __init__(self, sheet, section, view_port):
        self.Sheet = sheet
        self.Section = section
        self.ViewPort = view_port

    @property
    def SheetNumber(self):
        return self.Sheet.SheetNumber if self.IsFullName else self.Sheet.SheetNumber.split("-").pop()

    @property
    def ViewNumber(self):
        return self.Section.GetParamValueOrDefault("_Номер Вида на Листе")

    @property
    def IsFullName(self):
        return self.Section.GetParamValueOrDefault("_Полный Номер Листа")

    @property
    def DetailNumber(self):
        return "{} ({})".format(self.ViewNumber, self.SheetNumber)

    def UpdateParam(self):
        self.Section.SetParamValue(BuiltInParameter.VIEWPORT_DETAIL_NUMBER, self.DetailNumber)

    @staticmethod
    def GetViewSections(sheet):
        view_ports = [document.GetElement(element_id) for element_id in sheet.GetAllViewports()]

        view_sections = []
        for view_port in view_ports:
            section = document.GetElement(view_port.ViewId)
            if isinstance(section, ViewSection):
                view_sections.append(Section(sheet, section, view_port))

        return view_sections

    def __str__(self):
        return "{}".format(self.DetailNumber)

    def __eq__(self, other):
        if isinstance(other, Section):
            return self.Sheet.Id == other.Sheet.Id \
                   and self.DetailNumber == other.DetailNumber

        return False


def get_table_columns():
    return ["SectionId", "SectionName", "ViewPortId", "ViewPortName"]


def get_row_section(section):
    return [output.linkify(section.Section.Id), section.Section.Name,
            output.linkify(section.ViewPort.Id), section.ViewPort.Name]


def get_table_data(view_sections):
    return [ get_row_section(section) for section in view_sections ]

def show_error_sections(view_sections):
    error_sections = [section for section in view_sections if not section.ViewNumber]
    if error_sections:
        table_columns = get_table_columns()
        table_data = get_table_data(error_sections)
        show_alert("Виды с пустым атрибутом \"_Номер Вида на Листе\".", table_columns, table_data)


def show_duplicates_sections(view_sections):
    duplicates = groupby(view_sections, lambda x: x)

    duplicate_sections = []
    for section, sections in duplicates:
        sections = list(sections)
        sections_count = len(sections)

        if sections_count > 1:
            duplicate_sections.extend(sections)

    if duplicate_sections:
        table_columns = get_table_columns()
        table_data = get_table_data(duplicate_sections)
        show_alert("Найдено дублирование значений атрибута \"Номер вида\".", table_columns, table_data)


def show_alert(title, table_columns, table_data, exit_script=True):
    output.print_table(title=title, columns=table_columns, table_data=table_data)

    if exit_script:
        script.exit()


def update_view_number():
    view_sheets = FilteredElementCollector(document).OfClass(ViewSheet).ToElements()
    view_sections = [section for sheet in view_sheets
                     for section in Section.GetViewSections(sheet)]

    show_error_sections(view_sections)
    show_duplicates_sections(view_sections)

    with Transaction(document) as transaction:
        transaction.Start("Обновление номера вида")

        for section in view_sections:
            section.UpdateParam()

        transaction.Commit()


update_view_number()