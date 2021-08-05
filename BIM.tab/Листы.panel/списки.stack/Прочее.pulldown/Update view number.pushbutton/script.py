# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")

import dosymep.Revit
clr.ImportExtensions(dosymep.Revit)

from itertools import groupby

from System.Collections.Generic import List
from Autodesk.Revit.DB import *

from pyrevit import script
from pyrevit import forms

document = __revit__.ActiveUIDocument.Document


class ViewSection(object):
    def __init__(self, sheet, section):
        self.Sheet = sheet
        self.Section = section

    @property
    def SheetNumber(self):
        return self.Sheet.SheetNumber.split("-").pop()

    @property
    def ViewNumber(self):
        return self.Section.GetParamValueOrDefault("_Номер Вида на Листе")

    @property
    def DetailNumber(self):
        return "{} ({})".format(self.ViewNumber, self.SheetNumber)

    def UpdateParam(self):
        self.Section.SetParamValue(BuiltInParameter.VIEWPORT_DETAIL_NUMBER, self.DetailNumber)

    @staticmethod
    def GetViewSections(sheet):
        sections = [document.GetElement(element_id) for element_id in sheet.GetAllPlacedViews()]
        return [ViewSection(sheet, section) for section in sections]

    def __str__(self):
        return "{}".format(self.DetailNumber)


def show_error_sections(view_sections):
    error_sections = [section.Section.Name for section in view_sections if not section.ViewNumber]
    if error_sections:
        section_names = set(sorted(error_sections))
        show_alert("Виды с пустым атрибутом \"_Номер Вида на Листе\":", " - " + "\r\n - ".join(section_names))


def show_duplicates_sections(view_sections):
    duplicates = groupby(view_sections, lambda x: x.DetailNumber)
    duplicates = [(k, v) for k, v in duplicates if len(list(v)) > 1]
    if duplicates:
        message = ""
        for k, v in duplicates:
            print k, list(v)
            if v:
                message += k + ":"
                message += "\r\n" + "\r\n - ".join([s.Section.ViewName for s in v])

        show_alert("Найдено дублирование значений атрибута \"Номер вида\":", message)


def show_alert(message, sub_message, exit_script=True):
    forms.alert(message, sub_msg=sub_message, title="Предупреждение!", footer="dosymep (Обновление номера вида)", exitscript=exit_script)


def update_view_number():
    print "sss"
    view_sheets = FilteredElementCollector(document).OfClass(ViewSheet).ToElements()
    view_sections = [section for sheet in view_sheets
                     for section in ViewSection.GetViewSections(sheet)]

    show_error_sections(view_sections)
    show_duplicates_sections(view_sections)

    with Transaction(document) as transaction:
        transaction.Start("Обновление номера вида")

        for section in view_sections:
            section.UpdateParam()

        transaction.Commit()





update_view_number()
