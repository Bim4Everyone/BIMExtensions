# -*- coding: utf-8 -*-

import json
import datetime

from pyrevit import EXEC_PARAMS
from pyrevit.coreutils import envvars
from pyrevit.userconfig import user_config


TIME_PROPERTY = "time_sec"
PATH_NAME_PROPERTY = "path_name"

TIME_FORMAT = "%Y-%m-%d %H:%M:%S"
SYNC_DOC_TIME = "B4E_SYNC_DOCUMENT"


def syncing_document():
    args = EXEC_PARAMS.event_args

    data = {
        PATH_NAME_PROPERTY: args.Document.PathName,
        TIME_PROPERTY: datetime.datetime.now().strftime(TIME_FORMAT)
    }

    envvars.set_pyrevit_env_var(SYNC_DOC_TIME, json.dumps(data))


try:
    if user_config.log_trace.enable_sync_doc_time:
        syncing_document()
except:
    pass
