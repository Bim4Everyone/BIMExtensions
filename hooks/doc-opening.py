# -*- coding: utf-8 -*-

import json
import datetime

from pyrevit import EXEC_PARAMS
from pyrevit.coreutils import envvars

from dosymep_libs.simple_services import *

TIME_PROPERTY = "time_sec"
PATH_NAME_PROPERTY = "path_name"

TIME_FORMAT = "%Y-%m-%d %H:%M:%S"
OPEN_DOC_TIME = "B4E_OPEN_DOCUMENT"


def opening_document():
    args = EXEC_PARAMS.event_args

    data = {
        PATH_NAME_PROPERTY: args.PathName,
        TIME_PROPERTY: datetime.datetime.now().strftime(TIME_FORMAT)
    }

    envvars.set_pyrevit_env_var(OPEN_DOC_TIME, json.dumps(data))


opening_document()
