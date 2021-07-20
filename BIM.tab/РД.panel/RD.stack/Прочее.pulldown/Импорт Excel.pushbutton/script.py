# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
import clr
clr.AddReference('System')
clr.AddReference('System.Drawing')
clr.AddReference('System.IO')
clr.AddReference("System.Windows.Forms")
clr.AddReference("EPPlus")
from System.IO import FileInfo
from System.Windows.Forms import MessageBox, SaveFileDialog, DialogResult, FolderBrowserDialog 
from System.Drawing import Color as C_Color
from System.Collections.Generic import List 
from Autodesk.Revit.DB import Parameter, SectionType, ViewScheduleExportOptions, HorizontalAlignmentStyle, VerticalAlignmentStyle, ViewSchedule, FilteredElementCollector, SharedParameterElement
from OfficeOpenXml import ExcelPackage, Style
from pyrevit import forms

def FilterString(obj):
	res = obj
	unacceptableSymbols = ['/', '\\', ':', '*', '<', '>', '|']
	for symbol in unacceptableSymbols:
		temp = res.split(symbol)
		res = "".join(temp)
	
	return res

class TabelsConverter(object):
	def __init__(self, view):
		self.__view = view
		self.__data = view.GetTableData()
		self.__currentRow = 1
		self.__currentColumn = 1
		self.package = self.__createPackage()
		self.__currentWorksheet = self.package.Workbook.Worksheets.Add(FilterString(view.Name))
		self.__exportToExcel()

	def __createPackage(self):
		return ExcelPackage()
	
	def __exportToExcel(self):
		self.__alignCells()
		self.__exportSection(SectionType.Header)
		self.__exportSection(SectionType.Body)
	
	def __alignCells(self):
		sectionData = self.__data.GetSectionData(SectionType.Body)
		numberOfColumns = sectionData.NumberOfColumns
		for columnNumber in range(numberOfColumns):
			columnWidth = sectionData.GetColumnWidthInPixels(columnNumber) / 6
			self.__currentWorksheet.Column(columnNumber + 1).Width = columnWidth

	def __exportSection(self, sectionType):
		sectionData = self.__data.GetSectionData(sectionType)

		numberOfRows = sectionData.NumberOfRows
		numberOfColumns = sectionData.NumberOfColumns
		firstRow = sectionData.FirstRowNumber 
		firstColumn = sectionData.FirstColumnNumber

		for rowNumber in range(firstRow, firstRow + numberOfRows):
			rowHeight = sectionData.GetRowHeightInPixels(rowNumber)
			self.__currentWorksheet.Row(self.__currentRow).Height = rowHeight
			for columnNumber in range(firstColumn, firstColumn + numberOfColumns):
				self.__exportCell(rowNumber, columnNumber, sectionType)
				self.__currentColumn += 1
			self.__currentRow += 1
			self.__currentColumn = 1

	def __exportCell(self, row, column, sectionType):
		sectionData = self.__data.GetSectionData(sectionType)
		mergetCell = sectionData.GetMergedCell(row, column)
		if row == mergetCell.Top and column == mergetCell.Left:
			text = self.__view.GetCellText(sectionType, row, column)
			

			cellStyle = sectionData.GetTableCellStyle(row, column)

			excelCell = self.__currentWorksheet.Cells[self.__currentRow, self.__currentColumn]

			excelCell.Value = text

			splitText = text.split(",")
			if len(splitText) < 3:
				# print "may digit"
				if all(map(lambda x: x.isdigit(), splitText)):
					excelCell.Value = float(".".join(splitText))
					if len(splitText) == 1:
						excelCell.Style.Numberformat.Format = "0"
					else:
						excelCell.Style.Numberformat.Format = "0." + "0" * len(splitText[1])
					# print "DIGIT!!"

			excelCell.Style.Font.Bold = cellStyle.IsFontBold
			excelCell.Style.Font.Italic = cellStyle.IsFontItalic
			excelCell.Style.Font.UnderLine = cellStyle.IsFontUnderline
			excelCell.Style.Font.Size = cellStyle.TextSize
			excelCell.Style.Font.Name = cellStyle.FontName
			excelCell.Style.WrapText = True

			horizontalAlignment = cellStyle.FontHorizontalAlignment
			if horizontalAlignment == HorizontalAlignmentStyle.Center:
				excelCell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Center
			elif horizontalAlignment == HorizontalAlignmentStyle.Left:
				excelCell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Left
			else:
				excelCell.Style.HorizontalAlignment = Style.ExcelHorizontalAlignment.Right

			verticalAlignment = cellStyle.FontVerticalAlignment 
			if verticalAlignment == VerticalAlignmentStyle.Top:
				excelCell.Style.VerticalAlignment = Style.ExcelVerticalAlignment.Top
			elif verticalAlignment == VerticalAlignmentStyle.Middle:
				excelCell.Style.VerticalAlignment = Style.ExcelVerticalAlignment.Center
			else:
				excelCell.Style.VerticalAlignment = Style.ExcelVerticalAlignment.Bottom
			
			# param = self.__view.Document.GetElement(sectionData.GetCellParamId(row,
			# column))
			# if param:
			# 	if isinstance(param, Parameter):
			# 		print "-- {} - {}".format(param.Definition.Name, param.StorageType)
			# 	else:
			# 		definition = param.GetDefinition()
			# 		print "*****{}*".format(definition.Name)
			rowRange = mergetCell.Bottom - row
			columnRange = mergetCell.Right - column
			if rowRange > 0 or columnRange > 0:
				# print "Merge {},{} - {},{}".format(rowNumber, columnNumber,
				# mergetCell.Bottom, mergetCell.Right)
				self.__currentWorksheet.Cells[self.__currentRow, self.__currentColumn, self.__currentRow + rowRange, self.__currentColumn + columnRange].Merge = True


class CheckBoxOption:
	def __init__(self, obj):
		self.state = False
		self.name = obj.Name
		self.obj = obj

	def __nonzero__(self):
		return self.state

	def __str__(self):
		return self.name

__title__ = 'Выгрузка спецификаций в Excel'


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
view = __revit__.ActiveUIDocument.ActiveGraphicalView 
view = doc.ActiveView

viewSchedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
items = [CheckBoxOption(x) for x in viewSchedules]
items = sorted(items, key=lambda item: item.name)
res = forms.SelectFromList.show(items, button_name='Выбрать', multiselect=True)
if res is None:
	raise SystemExit(1)

# print len(items)
# print len([x for x in res if x.state])
# saveFileDialog = SaveFileDialog()
# saveFileDialog.Filter = "xlsx files (*.xlsx)|*.xlsx|All files (*.*)|*.*"
# saveFileDialog.FilterIndex = 1
# saveFileDialog.RestoreDirectory = True
folderBrowserDialog = FolderBrowserDialog()
folderBrowserDialog.Description = "Выберите папку для сохранения спецификаций."
folderBrowserDialog.ShowNewFolderButton = True

# if saveFileDialog.ShowDialog() == DialogResult.OK:
if folderBrowserDialog.ShowDialog() == DialogResult.OK:
	folderName = folderBrowserDialog.SelectedPath
	# fileName = saveFileDialog.FileName
	# fileInfo = FileInfo(fileName)

	for item in res:
		fileInfo = FileInfo(folderName + "\\" + FilterString(item.name) + ".xlsx")

		converter = TabelsConverter(item.obj)
		package = converter.package
		try:
			package.SaveAs(fileInfo)
			package.Save()
		except:
			print "Не удалось сохранить спецификацию '{}'".format(item.name)
		package.Dispose()