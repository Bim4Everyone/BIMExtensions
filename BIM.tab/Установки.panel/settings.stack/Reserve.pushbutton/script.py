# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
import clr
clr.AddReference('System')
clr.AddReference('System.IO')
clr.AddReference("System.Windows.Forms")
from System.IO import FileInfo
from System.Windows.Forms import MessageBox
from System.Collections.Generic import List
from Autodesk.Revit.DB import CurveArray

import clr
clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import MessageBox
import os.path as op
import datetime
import shutil

from Autodesk.Revit.DB import ModelPathUtils, BasicFileInfo, WorksharingSaveAsOptions
from pySpeech.Forms import InputFormText

class DocumentReserver:
	def __init__(self, document):
		self._document = document
		self._reserveDirectory = self._getReserveDocumentDirectory()
		if op.isdir(self._reserveDirectory):
			self._reservedName = self._getReserveDocumentName()
			self._refreshReservedName()
			centralDocumentPath = self._getCentralDocumentPath()
			reservePath = op.join(self._reserveDirectory, self._reservedName)
			shutil.copy2(centralDocumentPath, reservePath)
		else:
			MessageBox.Show('Структура проекта не соответствует стандарту!')

	def _refreshReservedName(self):
		name = op.join(self._reserveDirectory, self._reservedName)
		splitedName = name[:-4]
		suffix = InputFormText.show([], title='Введите Суффикс', button_name='Ок', width=250, height=130)
		if not suffix is None:
			name = splitedName + "_{}.rvt".format(suffix)

		start = 1
		while op.isfile(name):
			name = splitedName + "_{}.rvt".format(start)
			start += 1
		self._reservedName = op.basename(name)
	
	def _getCentralDocumentPath(self):
		path = self._document.GetWorksharingCentralModelPath()
		centralModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(path)
		centralModelPath = BasicFileInfo.Extract(doc.PathName).CentralPath
		return centralModelPath

	def _getReserveDocumentDirectory(self):
		centralModelPath = self._getCentralDocumentPath()
		centralModelPathDirectory = op.dirname(centralModelPath)
		rootProjectDirectory = op.dirname(centralModelPathDirectory)
		newPathDirectory = op.join(rootProjectDirectory, "4 - Резервные копии")
		return newPathDirectory

	def _getReserveDocumentName(self):
		centralModelPath = self._getCentralDocumentPath()
		baseName = op.basename(centralModelPath)
		today = datetime.date.today()
		datedName = today.strftime("%Y-%m-%d") + "_" + baseName
		return datedName


doc = __revit__.ActiveUIDocument.Document
if doc.IsWorkshared:
	DocumentReserver(doc)
else:
	MessageBox.Show('Открыт отсоединенный файл!')


