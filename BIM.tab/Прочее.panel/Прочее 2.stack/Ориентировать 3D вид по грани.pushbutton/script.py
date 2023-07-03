# coding=utf-8
from pyrevit import HOST_APP
from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


def reorient():
    face = revit.pick_face()

    if face:
        with revit.Transaction('Orient to Selected Face'):
            # calculate normal
            if HOST_APP.is_newer_than(2015):
                normal_vec = face.ComputeNormal(DB.UV(0, 0))
            else:
                normal_vec = face.Normal

            # create base plane for sketchplane
            if HOST_APP.is_newer_than(2016):
                base_plane = \
                    DB.Plane.CreateByNormalAndOrigin(normal_vec, face.Origin)
            else:
                base_plane = DB.Plane(normal_vec, face.Origin)

            # now that we have the base_plane and normal_vec
            # let's create the sketchplane
            sp = DB.SketchPlane.Create(revit.doc, base_plane)

            # orient the 3D view looking at the sketchplane
            revit.active_view.OrientTo(normal_vec.Negate())
            # set the sketchplane to active
            revit.uidoc.ActiveView.SketchPlane = sp

        revit.uidoc.RefreshActiveView()


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    curview = revit.active_view

    if isinstance(curview, DB.View3D):
        reorient()
    else:
        forms.alert('Должен быть открыт 3D вид.')


script_execute()