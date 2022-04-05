# -*- coding: utf-8 -*-

from pyrevit import EXEC_PARAMS
from dosymep_libs.simple_services import *

logger = get_logger_service()
logger.Debug("Инициализация платформы.")

show_notification_service("Bim4Everyone", "Инициализация платформы прошла успешно.")
