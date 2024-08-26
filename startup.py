# -*- coding: utf-8 -*-

import os
import sys
from pyrevit.versionmgr import updater
import LibGit2Sharp as libgit
from pyrevit import coreutils
import tempfile


def update_extensions():
    revit_count = coreutils.get_revit_instance_count()
    temp_dir_exist = os.path.isdir(os.path.join(tempfile.gettempdir(), "Bim4Everyone", "dosymep"))
    if revit_count == 1 and not temp_dir_exist:
        for repo_info in updater.get_all_extension_repos():
            try:
                if updater.has_pending_updates(repo_info):
                    # сбрасываем репозиторий в исходное состояние
                    repo_info.repo.Reset(libgit.ResetMode.Hard, repo_info.repo.Head.Tip)
                    repo_info.repo.RemoveUntrackedFiles()

                    # пытаемся обновится
                    updater.update_repo(repo_info)
            except Exception:
                pass


update_extensions()


# при запуске pyrevit
# не указывает путь до библиотеки
directory = os.getenv('APPDATA')
bim4everyone_lib = r"pyRevit\Extensions\BIM4Everyone.lib"
bim4everyone_lib_path = os.path.join(directory, bim4everyone_lib)
sys.path.append(bim4everyone_lib_path)

import dosymep_libs
dosymep_libs.load_assemblies()

import devexpress_libs
devexpress_libs.load_assemblies()

import clr
clr.AddReference('dosymep.Revit.dll')
clr.AddReference('dosymep.Bim4Everyone.dll')

clr.AddReference('dosymep.Xpf.Core.dll')
clr.AddReference('dosymep.SimpleServices.dll')

clr.AddReference('Serilog.dll')
clr.AddReference('Serilog.Sinks.File.dll')

clr.AddReference('Serilog.Sinks.Autodesk.Revit.dll')
clr.AddReference('Serilog.Enrichers.Autodesk.Revit.dll')

clr.AddReference('DevExpress.Xpf.Core.v21.2.dll')
clr.AddReference('DevExpress.Data.Desktop.v21.2.dll')

clr.AddReference('DevExpress.Dialogs.v21.2.Core.dll')

from pyrevit import HOST_APP
from pyrevit.userconfig import user_config

from DevExpress.Xpf.Core import *

from dosymep_libs.simple_services import *
from dosymep.Bim4Everyone.Schedules import SchedulesConfig
from dosymep.Bim4Everyone.KeySchedules import KeySchedulesConfig
from dosymep.Bim4Everyone.SystemParams import SystemParamsConfig
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig


def get_config_path(section, option):
    try:
        if user_config.has_section(section):
            return user_config.get_section(section).get_option(option)
    except:
        pass


def load_platform_settings():
    shared_params_path = get_config_path("PlatformSettings", "SharedParamsPath")
    project_params_path = get_config_path("PlatformSettings", "ProjectParamsPath")

    SchedulesConfig.LoadInstance("")
    KeySchedulesConfig.LoadInstance("")
    SystemParamsConfig.LoadInstance(HOST_APP.language)
    SharedParamsConfig.LoadInstance(shared_params_path)
    ProjectParamsConfig.LoadInstance(project_params_path)

    # Удаляем стиль DevExpress
    # (на некоторых окнах портится layout)
    ApplicationThemeHelper.ApplicationThemeName = Theme.NoneName

    ServicesProvider.LoadInstance(__revit__)


load_platform_settings()
