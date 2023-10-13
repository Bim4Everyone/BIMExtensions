# -*- coding: utf-8 -*-

from System.IO import File
from System.Text import Encoding
from System.Diagnostics import Process

from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

document = __revit__.ActiveUIDocument.Document


def generate_column_headers(sizeTable):
    columns = [""]
    for column_index in range(1, sizeTable.NumberOfColumns):
        column_header = sizeTable.GetColumnHeader(column_index)

        unit_type = "OTHER"
        if HOST_APP.is_newer_than(2021):
            if UnitUtils.IsMeasurableSpec(column_header.GetSpecTypeId()):
                unit_type = UnitUtils.GetTypeCatalogStringForSpec(column_header.GetSpecTypeId())
        else:
            if column_header.UnitType != UnitType.UT_Undefined:
                unit_type = UnitUtils.GetTypeCatalogString(column_header.UnitType)

        display_unit_type = ""
        if HOST_APP.is_newer_than(2021):
            if UnitUtils.IsUnit(column_header.GetUnitTypeId()):
                display_unit_type = UnitUtils.GetTypeCatalogStringForUnit(column_header.GetUnitTypeId())
        else:
            if column_header.DisplayUnitType != DisplayUnitType.DUT_UNDEFINED:
                display_unit_type = UnitUtils.GetTypeCatalogString(column_header.DisplayUnitType)

        columns.append("{}##{}##{}".format(column_header.Name, unit_type, display_unit_type))

    return columns


def generate_table(family, table_name):
    size_table = family.GetSizeTable(table_name)

    result = []
    result.append(";".join(generate_column_headers(size_table)))

    for rowIndex in range(0, size_table.NumberOfRows):
        columns = []

        for columnIndex in range(0, size_table.NumberOfColumns):
            columns.append(size_table.AsValueString(rowIndex, columnIndex))

        result.append(";".join(columns))

    return result


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    families = FilteredElementCollector(document).OfClass(Family).ToElements()
    families = [FamilyElement(document, family) for family in families]
    families = [family for family in families if family.HasSizeTables]

    family = forms.SelectFromList.show(families, button_name='Выбрать')
    if not family:
        script.exit()

    table_name = forms.SelectFromList.show(family.SizeTableNames, button_name='Выбрать')
    if not table_name:
        script.exit()

    result = generate_table(family, table_name)
    file_name = forms.save_file(files_filter="csv files (*.csv)|*.csv", default_name=table_name)
    if file_name:
        File.WriteAllText(file_name, "\r\n".join(result), Encoding.GetEncoding(1251))
        Process.Start(file_name)


class FamilyElement:
    def __init__(self, document, family):
        self.Id = family.Id

        self.Name = family.Name
        if not self.Name:
            self.Name = "Отсутствует наименование"

        self.SizeTableNames = None
        self.TableManager = FamilySizeTableManager.GetFamilySizeTableManager(document, family.Id)
        if self.TableManager:
            self.SizeTableNames = self.TableManager.GetAllSizeTableNames()

    @property
    def SizeTableCount(self):
        if self.SizeTableNames:
            return self.SizeTableNames.Count

        return 0

    @property
    def HasSizeTables(self):
        return self.SizeTableCount > 0

    def GetSizeTable(self, tableName):
        return self.TableManager.GetSizeTable(tableName)

    def __str__(self):
        return "{} [{}]".format(self.Name, self.SizeTableCount)


script_execute()
