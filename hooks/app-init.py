# -*- coding: utf-8 -*-

from pyrevit import EXEC_PARAMS
from pyrevit.versionmgr import updater

from dosymep_libs.simple_services import *

logger = get_logger_service()
logger.Debug("Инициализация платформы.")


def check_updates():
    status_update = None
    for repo_info in updater.get_all_extension_repos():
        if updater.has_pending_updates(repo_info):
            status_update = True
            logger.Debug("Ошибка обновления расширения: \"{}\".".format(repo_info))

    return status_update


if check_updates():
    show_notification_service("Bim4Everyone",
                              "Инициализация платформы прошла с ошибкой, пожалуйста переустановите платформу.")
else:
    show_notification_service("Bim4Everyone",
                              "Инициализация платформы прошла успешно.")
