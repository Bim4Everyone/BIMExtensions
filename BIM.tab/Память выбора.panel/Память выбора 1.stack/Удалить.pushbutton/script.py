import os
import os.path as op
import pickle as pl
import clr

clr.AddReference("dosymep.Revit.dll")

from pyrevit import revit
from pyrevit import script
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
        prevsel = pl.load(f)
        newsel = prevsel.difference(selected_ids)
        f.close()
        f = open(datafile, 'w')
        pl.dump(newsel, f)
        f.close()
    except Exception:
        f = open(datafile, 'w')
        pl.dump([], f)
        f.close()


script_execute()
