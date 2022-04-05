# -*- coding: utf-8 -*-

from System.Collections.Generic import *

from pyrevit import EXEC_PARAMS
from dosymep_libs.simple_services import *

document = EXEC_PARAMS.event_args.Document
dictionary = {
    "Title": document.Title,
    "PathName": document.PathName,
    "IsWorkshared": document.IsWorkshared,
    "IsModelInCloud": document.IsModelInCloud,
    "IsFamilyDocument": document.IsFamilyDocument,
    "DisplayUnitSystem": document.DisplayUnitSystem
}

logger = get_logger_service()
logger.Debug("Синхронизация документа {@document}.", Dictionary[str, object](dictionary))