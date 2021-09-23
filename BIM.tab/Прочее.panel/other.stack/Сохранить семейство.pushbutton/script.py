# -*- coding: utf-8 -*-

import os
import os.path as op
import tempfile

from Autodesk.Revit.DB import *

uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
app = doc.Application

file_path = doc.PathName
file_name = op.basename(file_path)

temp_file_path = op.join(tempfile.gettempdir(), file_name)

if doc.IsFamilyDocument:
	saveOptions = SaveAsOptions()
	saveOptions.Compact = False
	saveOptions.OverwriteExistingFile = True

	doc.SaveAs(temp_file_path, saveOptions)
	try:
		doc.SaveAs(file_path, saveOptions)
	finally:
		os.remove(temp_file_path)