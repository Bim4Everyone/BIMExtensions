import pickle

from pyrevit import EXEC_PARAMS
from pyrevit import revit, DB

from dosymep_libs.bim4everyone import *


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    try:
        selection = revit.get_selection()
        datafile = script.get_document_data_file("SelList", "pym")

        f = open(datafile, 'r')
        current_selection = pickle.load(f)
        f.close()

        element_ids = []
        for elid in current_selection:
            element_ids.append(DB.ElementId(int(elid)))

        selection.set_to(element_ids)
    except Exception:
        pass


script_execute()
