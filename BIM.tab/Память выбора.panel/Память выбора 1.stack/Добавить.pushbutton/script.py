import pickle
import clr

clr.AddReference("dosymep.Revit.dll")

from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

import dosymep

clr.ImportExtensions(dosymep.Revit)


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    datafile = script.get_document_data_file("SelList", "pym")

    selection = revit.get_selection()
    selected_ids = {str(elid.GetIdValue()) for elid in selection.element_ids}

    try:
        f = open(datafile, 'r')
        prevsel = pickle.load(f)
        new_selection = prevsel.union(selected_ids)
        f.close()
    except Exception:
        new_selection = selected_ids

    f = open(datafile, 'w')
    pickle.dump(new_selection, f)
    f.close()


script_execute()
