# -*- coding: utf-8 -*-
import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from System.IO import Path
from Autodesk.Revit.DB import *
from Autodesk.Revit.Exceptions import *

from pyrevit import forms
from pyrevit import script
from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

document = __revit__.ActiveUIDocument.Document


def reload_existing_links(selected_files, progress, cancellation_token):
    linked_files = FilteredElementCollector(document)\
        .OfCategory(BuiltInCategory.OST_RvtLinks)\
        .WhereElementIsElementType()\
        .ToElements()

    link_files = [(link_file, link_file.GetParamValue(BuiltInParameter.ALL_MODEL_TYPE_NAME)) for link_file in linked_files]
    link_files = [(link_file, selected_files.pop(link_name.lower(), None)) for (link_file, link_name) in link_files]
    link_files = [(link_file, link_file_path) for link_file, link_file_path in link_files if link_file_path]

    count = float(len(link_files))
    for i, (link_file, link_file_path) in enumerate(link_files):
        progress.Report(int(round((i + 1) / count * 100)))
        cancellation_token.ThrowIfCancellationRequested()

        link_file_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(link_file_path)
        link_file.LoadFrom(link_file_model_path, WorksetConfiguration())
            

def create_revit_links(selected_files, progress, cancellation_token):
    error_list = []
    error_version_list = []

    with revit.Transaction("BIM: Связывание файлов", log_errors=False):
        count = float(len(selected_files.items()))
        for i, (file_name, file_name_path) in enumerate(selected_files.items()):
            progress.Report(int(round((i + 1) / count * 100)))
            cancellation_token.ThrowIfCancellationRequested()

            try:
                link_file = ModelPathUtils.ConvertUserVisiblePathToModelPath(file_name_path)

                link_options = RevitLinkOptions(True)
                link_load_result = RevitLinkType.Create(document, link_file, link_options)

                revit_link_instance = RevitLinkInstance.Create(document, link_load_result.ElementId, ImportPlacement.Shared)

                revit_link_type = document.GetElement(revit_link_instance.GetTypeId())
                revit_link_type.Parameter[BuiltInParameter.WALL_ATTR_ROOM_BOUNDING].Set(1)
            except InvalidOperationException:
                error_list.append(Path.GetFileName(file_name_path))
            except Exception:
                error_version_list.append(Path.GetFileName(file_name_path))

    error_message = None
    if error_list:
        error_message = "Файлы с другой системой координат:\r\n - " + "\r\n - ".join(error_list)

    if error_version_list:
        error_message = error_message if error_message else "" + "\r\nФайлы с созданные в другой версии Revit:\r\n - " + "\r\n - ".join(error_version_list)

    return error_message



@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selected_files = forms.pick_file(files_filter="Revit files (*.rvt)|*.rvt", multi_file=True, title="Выберите Revit файлы")
    if not selected_files:
        script.exit()

    selected_file_names = [ Path.GetFileName(value).lower() for value in list(selected_files) ]
    selected_files = dict(zip(selected_file_names, selected_files))

    error_message = None
    with get_progress_dialog_service() as progress_dialog:
        progress_dialog.MaxValue = 100
        progress_dialog.Show()

        progress = progress_dialog.CreateProgress()
        cancellation_token = progress_dialog.CreateCancellationToken()

        # Перезагружаем существующие связи
        progress_dialog.DisplayTitleFormat = "Обновление связей [{0}\\{1}] ..."
        reload_existing_links(selected_files, progress, cancellation_token)

        # Добавляем оставшиеся связи
        progress_dialog.DisplayTitleFormat = "Связывание файлов [{0}\\{1}] ..."
        error_message = create_revit_links(selected_files, progress, cancellation_token)

    if error_message:
        forms.alert(error_message.strip("\r\n"), title="Предупреждение", exitscript=True)


script_execute()