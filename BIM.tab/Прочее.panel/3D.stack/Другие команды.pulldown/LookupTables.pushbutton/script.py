# -*- coding: utf-8 -*-

from System.IO import File
from System.Text import Encoding
from System.Diagnostics import Process

from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script
from pyrevit import HOST_APP


document = __revit__.ActiveUIDocument.Document


def generate_column_headers(sizeTable):
    columns = [ "" ]
    for columnIndex in range(1, sizeTable.NumberOfColumns):
        columnHeader = sizeTable.GetColumnHeader(columnIndex)

        unitType = "OTHER"
        if HOST_APP.is_newer_than(2021):
            if UnitUtils.IsUnit(columnHeader.GetUnitTypeId()):
                unitType = UnitUtils.GetTypeCatalogStringForUnit(columnHeader.GetUnitTypeId())
        else:
            if columnHeader.UnitType != UnitType.UT_Undefined:
                unitType = UnitUtils.GetTypeCatalogString(columnHeader.UnitType)

        displayUnitType = ""
        if HOST_APP.is_newer_than(2021):
            if UnitUtils.IsMeasurableSpec(columnHeader.GetSpecTypeId()):
                displayUnitType = UnitUtils.GetTypeCatalogStringForSpec(columnHeader.GetSpecTypeId())
        else:
            if columnHeader.DisplayUnitType != DisplayUnitType.DUT_UNDEFINED:
                displayUnitType = UnitUtils.GetTypeCatalogString(columnHeader.DisplayUnitType)

        columns.append("{}##{}##{}".format(columnHeader.Name, unitType, displayUnitType))
    
    return columns

def generate_table(family, table_name):
    sizeTable = family.GetSizeTable(table_name)

    result = []
    result.append(";".join(generate_column_headers(sizeTable)))    

    for rowIndex in range(0, sizeTable.NumberOfRows):
        columns = []
    
        for columnIndex in range(0, sizeTable.NumberOfColumns):
            columns.append(sizeTable.AsValueString(rowIndex, columnIndex))
    
        result.append(";".join(columns))
    
    return result;

def excecute_script():
    familys = FilteredElementCollector(document).OfClass(Family).ToElements()
    familys = [ FamilyElement(document, family) for family in familys ]
    familys = [ family for family in familys if family.HasSizeTables ]

    family = forms.SelectFromList.show(familys, button_name='Выбрать')
    if not family:
        script.exit()

    tableName = forms.SelectFromList.show(family.SizeTableNames, button_name='Выбрать')
    if not tableName:
        script.exit()

    result = generate_table(family, tableName)        
    fileName = forms.save_file(files_filter="csv files (*.csv)|*.csv", default_name=tableName)
    if fileName:
        File.WriteAllText(fileName, "\r\n".join(result), Encoding.GetEncoding(1251))
        Process.Start(fileName)

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

#start script
excecute_script()