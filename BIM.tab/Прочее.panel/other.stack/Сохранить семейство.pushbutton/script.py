# -*- coding: utf-8 -*-

import os
import os.path as op
import tempfile
import re

from Autodesk.Revit.DB import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
app = doc.Application

file_path = doc.PathName
file_name = op.basename(file_path)

temp_file_path = op.join(tempfile.gettempdir(), file_name)


def remove_backups():
	# удаляем свой бекап файла ревита
	os.remove(temp_file_path)

	# удаляем бекапы созданные ревитом
	dir_name = op.dirname(file_path)
	backup_files = [op.join(dir_name, f) for f in os.listdir(dir_name)
					if f.startswith(doc.Title) and f != file_name and re.findall(r".*\.\d\d\d\d\..*", f) ]
	backup_files = [f for f in backup_files if op.isfile(f) ]
	for backup_file in backup_files:
		os.remove(backup_file)


if doc.IsFamilyDocument:
	saveOptions = SaveAsOptions()
	saveOptions.Compact = False
	saveOptions.OverwriteExistingFile = True

	doc.Save()
	doc.SaveAs(temp_file_path, saveOptions)
	try:
		doc.SaveAs(file_path, saveOptions)
	finally:
		remove_backups()