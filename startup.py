# -*- coding: utf-8 -*-

import os
import sys

# при запуске pyrevit
# не указывает путь до библиотеки
directory = os.getenv('APPDATA')
bim4everyone_lib = r"pyRevit\Extensions\BIM4Everyone.lib"
bim4everyone_lib_path = os.path.join(directory, bim4everyone_lib)
sys.path.append(bim4everyone_lib_path)

import dosymep_libs
dosymep_libs.load_assemblies()

import clr
clr.AddReference('dosymep.Revit.dll')
clr.AddReference('dosymep.Bim4Everyone.dll')

from pyrevit import HOST_APP
from pyrevit.userconfig import user_config

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


load_platform_settings()
