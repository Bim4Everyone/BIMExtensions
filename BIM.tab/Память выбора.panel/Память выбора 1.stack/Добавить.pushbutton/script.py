import pickle

import clr
clr.AddReference("dosymep.Revit.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)

from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    datafile = script.get_document_data_file("SelList", "pym")

    selection = revit.get_selection()
    selected_ids = {str(element_id.GetIdValue()) for element_id in selection.element_ids}

    try:
        with open(datafile, 'r') as f:
            prev_selection = pickle.load(f)
            new_selection = prev_selection.union(selected_ids)
    except Exception:
        new_selection = selected_ids

    with open(datafile, 'w') as f:
        pickle.dump(new_selection, f)


script_execute()
