import pickle

import clr
clr.AddReference("dosymep.Revit.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)

from pyrevit import revit
from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

from dosymep_libs.bim4everyone import log_plugin


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    datafile = script.get_document_data_file("SelList", "pym")

    selection = revit.get_selection()
    selected_ids = {str(elid.GetIdValue()) for elid in selection.element_ids}

    with open(datafile, 'w') as f:
        pickle.dump(selected_ids, f)


script_execute()
