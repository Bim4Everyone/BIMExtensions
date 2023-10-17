import pickle

from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS
from pyrevit import revit, DB

from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    try:
        selection = revit.get_selection()
        datafile = script.get_document_data_file("SelList", "pym")

        with open(datafile, 'r') as f:
            current_selection = pickle.load(f)

        element_ids = []
        for element_id in current_selection:
            if HOST_APP.is_older_than(2024):
                element_ids.append(DB.ElementId(int(element_id)))
            else:
                element_ids.append(DB.ElementId(long(element_id)))

        selection.set_to(element_ids)
    except Exception:
        pass


script_execute()
