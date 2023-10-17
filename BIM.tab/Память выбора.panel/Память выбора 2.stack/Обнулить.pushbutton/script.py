import pickle

from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    datafile = script.get_document_data_file("SelList", "pym")

    with open(datafile, 'wb') as f:
        pickle.dump([], f)


script_execute()
