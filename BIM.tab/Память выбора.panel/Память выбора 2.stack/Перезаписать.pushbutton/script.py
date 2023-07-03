import pickle

from pyrevit import script
from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import log_plugin


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    datafile = script.get_document_data_file("SelList", "pym")

    selection = revit.get_selection()
    selected_ids = {str(elid.IntegerValue) for elid in selection.element_ids}

    f = open(datafile, 'w')
    pickle.dump(selected_ids, f)
    f.close()


script_execute()
