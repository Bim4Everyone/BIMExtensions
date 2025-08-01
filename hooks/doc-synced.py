# -*- coding: utf-8 -*-

import json
from collections import OrderedDict
from datetime import datetime

from pyrevit import EXEC_PARAMS
from pyrevit.coreutils import envvars
from pyrevit.userconfig import user_config

from Autodesk.Revit.DB import *
from Autodesk.Revit.DB.Events import *
from System.Collections.Generic import Dictionary

from dosymep_libs.simple_services import *

TIME_PROPERTY = "time_sec"
PATH_NAME_PROPERTY = "path_name"

TIME_FORMAT = "%Y-%m-%d %H:%M:%S"
SYNC_DOC_TIME = "B4E_SYNC_DOCUMENT"


def get_visible_path(document):
    if document.IsWorkshared:
        model_path = document.GetWorksharingCentralModelPath()
        return ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)
    elif hasattr(document, "IsCloudModel") and document.IsCloudModel:
        model_path = document.GetCloudModelPath()
        return ModelPathUtils.ConvertModelPathToUserVisiblePath(model_path)

    return document.PathName


def synced_document():
    logger = get_logger_service()
    args = EXEC_PARAMS.event_args

    json_value = envvars.get_pyrevit_env_var(SYNC_DOC_TIME)
    try:
        data = json.loads(str(json_value))

        if args.Status != RevitAPIEventStatus.Succeeded:
            return

        if data[PATH_NAME_PROPERTY] != args.Document.PathName:
            return

        delta = datetime.now() - datetime.strptime(data[TIME_PROPERTY], TIME_FORMAT)
        logger.Information("Время синхронизации документа: {@DocSyncTime}", delta.total_seconds())
    finally:
        envvars.set_pyrevit_env_var(SYNC_DOC_TIME, None)


try:
    if user_config.log_trace.enable_sync_doc_time:
        synced_document()
except:
    pass
