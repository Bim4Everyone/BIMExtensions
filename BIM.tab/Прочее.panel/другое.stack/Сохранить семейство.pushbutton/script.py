# -*- coding: utf-8 -*-

import os
import os.path as op
import tempfile
import re

from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

from Autodesk.Revit.DB import *

document = __revit__.ActiveUIDocument.Document

file_path = document.PathName
file_name = op.basename(file_path)

temp_file_path = op.join(tempfile.gettempdir(), file_name)


def remove_backups():
	# удаляем свой бекап файла ревита
	os.remove(temp_file_path)

	# удаляем бекапы созданные ревитом
	dir_name = op.dirname(file_path)
	backup_files = [op.join(dir_name, f) for f in os.listdir(dir_name)
					if f.startswith(document.Title) and f != file_name and re.findall(r".*\.\d\d\d\d\..*", f)]
	backup_files = [f for f in backup_files if op.isfile(f) ]
	for backup_file in backup_files:
		os.remove(backup_file)


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
	if not document.IsFamilyDocument:
		show_script_warning_notification("Сохранять разрешено только файлы семейств", exit_script=True)

	try:
		save_options = SaveAsOptions()
		save_options.Compact = False
		save_options.OverwriteExistingFile = True

		document.Save()
		document.SaveAs(temp_file_path, save_options)
		try:
			document.SaveAs(file_path, save_options)
		finally:
			remove_backups()
	except:
		show_fatal_script_notification()
		raise

	show_executed_script_notification()


script_execute()