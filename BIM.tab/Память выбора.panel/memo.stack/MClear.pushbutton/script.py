import os
import os.path as op
import pickle as pl

from pyrevit import script
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import log_plugin


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    datafile = script.get_document_data_file("SelList", "pym")

    f = open(datafile, 'wb')
    pl.dump([], f)
    f.close()


script_execute()
