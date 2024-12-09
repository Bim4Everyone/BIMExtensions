# -*- coding: utf-8 -*-

import LibGit2Sharp as libgit
from System.Collections.Generic import Dictionary

from pyrevit.versionmgr import updater
from pyrevit.coreutils import envvars
from pyrevit.userconfig import user_config
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
    if user_config.auto_update \
            and not check_update_inprogress():
        set_autoupdate_inprogress(True)
        for repo_info in updater.get_all_extension_repos():
            try:
                if updater.has_pending_updates(repo_info):
                    logger.Warning("Репозиторий не был обновлен: \"{@RepoInfo}\"", to_dictionary(repo_info))
            except Exception:
                status_update = True
                logger.Warning("Ошибка обновления расширения: \"{@RepoInfo}\"", to_dictionary(repo_info))
        set_autoupdate_inprogress(False)

    return status_update

def check_update_inprogress():
    return envvars.get_pyrevit_env_var(envvars.AUTOUPDATING_ENVVAR)

def set_autoupdate_inprogress(state):
    envvars.set_pyrevit_env_var(envvars.AUTOUPDATING_ENVVAR, state)

if check_updates():
    log_trace("Инициализация платформы прошла с ошибкой")
    show_notification_service("Bim4Everyone",
                              "Инициализация платформы прошла с ошибкой, пожалуйста переустановите платформу.")
else:
    log_trace("Инициализация платформы прошла успешно")
    show_notification_service("Bim4Everyone",
                              "Инициализация платформы прошла успешно.")
