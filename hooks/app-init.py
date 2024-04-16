# -*- coding: utf-8 -*-

import LibGit2Sharp as libgit
from System.Collections.Generic import Dictionary

from pyrevit.versionmgr import updater
from dosymep_libs.simple_services import *

logger = get_logger_service()


def to_dictionary(repo_info):
    return Dictionary[str, object]({
        "name": repo_info.name,
        "branch": repo_info.branch,
        "directory": repo_info.directory,
        "head_name": repo_info.head_name,
        "last_commit_hash": repo_info.last_commit_hash,
    })


def log_trace(message):
    repo_infos = {}
    for extension in updater.get_all_extension_repos():
        repo_infos[extension.name] = to_dictionary(extension)

    logger.Information(message + ": {@Extensions}", Dictionary[str, object](repo_infos))


def check_updates():
    status_update = None
    for repo_info in updater.get_all_extension_repos():
        try:
            if updater.has_pending_updates(repo_info):
                logger.Warning("Репозиторий не на последнем коммите: \"{@RepoInfo}\"", to_dictionary(repo_info))
                # сбрасываем репозиторий в исходное состояние
                repo_info.repo.Reset(libgit.ResetMode.Hard, repo_info.repo.Head.Tip)
                repo_info.repo.RemoveUntrackedFiles()
                logger.Warning("Все изменения в локальном репозитории сброшены до текущего коммита: \"{@RepoInfo}\"",
                               to_dictionary(repo_info))
        except Exception:
            status_update = True
            logger.Warning("Ошибка обновления расширения: \"{@RepoInfo}\"", to_dictionary(repo_info))

    return status_update


if check_updates():
    log_trace("Инициализация платформы прошла с ошибкой")
    show_notification_service("Bim4Everyone",
                              "Инициализация платформы прошла с ошибкой, пожалуйста переустановите платформу.")
else:
    log_trace("Инициализация платформы прошла успешно")
    show_notification_service("Bim4Everyone",
                              "Инициализация платформы прошла успешно.")
