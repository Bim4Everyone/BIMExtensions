# -*- coding: utf-8 -*-

import LibGit2Sharp as libgit
from pyrevit.versionmgr import updater

from dosymep_libs.simple_services import *

logger = get_logger_service()
logger.Debug("Инициализация платформы.")


def check_updates():
    status_update = None
    for repo_info in updater.get_all_extension_repos():
        try:
            if updater.has_pending_updates(repo_info):
                # сбрасываем репозиторий в исходное состояние
                repo_info.repo.Reset(libgit.ResetMode.Hard, repo_info.repo.Head.Tip)
                repo_info.repo.RemoveUntrackedFiles()

                # пытаемся обновится
                updater.update_repo(repo_info)
        except Exception as ex:
            status_update = True
            logger.Debug(ex, "Ошибка обновления расширения: \"{}\".".format(repo_info))

    return status_update


if check_updates():
    show_notification_service("Bim4Everyone",
                              "Инициализация платформы прошла с ошибкой, пожалуйста переустановите платформу.")
else:
    show_notification_service("Bim4Everyone",
                              "Инициализация платформы прошла успешно.")
