import pickle
import clr

clr.AddReference("dosymep.Revit.dll")

from pyrevit import script
from pyrevit import revit
from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *

import dosymep

from dosymep_libs.bim4everyone import log_plugin

clr.ImportExtensions(dosymep.Revit)


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    datafile = script.get_document_data_file("SelList", "pym")

    selection = revit.get_selection()
    selected_ids = {str(elid.GetIdValue()) for elid in selection.element_ids}

    f = open(datafile, 'w')
    pickle.dump(selected_ids, f)
    f.close()


script_execute()
