# -*- coding: utf-8 -*-
import os
from os import walk
import os.path as op
import clr
clr.AddReference("System.Windows.Forms")
from pyrevit import EXTENSIONS_DEFAULT_DIR
from System.Windows.Forms import OpenFileDialog, DialogResult
from System.Windows.Forms import FolderBrowserDialog, DialogResult

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
app = doc.Application
import os
__title__ = 'Сохранить\nсемейство'
FilePath = doc.PathName
FileName = doc.Title
#FilePath = FilePath[:FilePath.rindex(FileName)]

TempFilePath = op.join(EXTENSIONS_DEFAULT_DIR, FileName)

if doc.IsFamilyDocument:
	doc.SaveAs(TempFilePath)
	os.remove(FilePath)
	doc.SaveAs(FilePath)
	os.remove(TempFilePath)

# fileContent = ''
# filePath = ''
# openFileDialog = OpenFileDialog()

# openFileDialog.InitialDirectory = "c:\\"
# openFileDialog.Filter = "rfa files (*.rfa)|*.rfa|All files (*.*)|*.*"
# openFileDialog.FilterIndex = 2
# openFileDialog.RestoreDirectory = True
# if (openFileDialog.ShowDialog() == DialogResult.OK):
# 	filePath = openFileDialog.FileName
# 	print filePath

	
# def reSave(FilePath, FileName):
	# fileName = op.join(FilePath, FileName)
	# doc = app.OpenDocumentFile(fileName)
	# doc.SaveAs(fileName+'_TEMP')
	# os.remove(fileName)
	# doc.SaveAs(fileName)
	# os.remove(fileName+'_TEMP')
	# try:
		# doc.Close(False)
	# except:
		# pass
	
	
# folderBrowserDialog = FolderBrowserDialog()
# folderBrowserDialog.Description = "Select the directory that you want to use as the default."
# folderBrowserDialog.ShowNewFolderButton = False
# result = folderBrowserDialog.ShowDialog()
# if result == DialogResult.OK:
	# folderName = folderBrowserDialog.SelectedPath

	
# for (dirpath, dirnames, filenames) in walk(folderName):
	# for filename in filenames:
		# suffix = filename.split('.')[-1]
		# if suffix == 'rfa':
			# reSave(dirpath,filename)
		