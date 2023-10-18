# coding=utf-8
"""Aligns the section box of the current 3D view to selected face."""
import math

from pyrevit.framework import Math
from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *
from dosymep_libs.simple_services import *

cur_view = revit.active_view


def orient_section_box(view):
    with forms.WarningBar(title="Выделите грань по которой требуется выравнять границы 3D вида"):
        face = revit.pick_face("Выделите грань по которой требуется выравнять границы 3D вида")
        if face is None:
            script.exit()

    box = view.GetSectionBox()
    face_norm = face.ComputeNormal(DB.UV(0, 0)).Normalize()
    face_angle = DB.XYZ.BasisX.AngleOnPlaneTo(face_norm, DB.XYZ.BasisZ)
    transform = box.Transform
    box_angle = transform.BasisX.AngleOnPlaneTo(transform.OfVector(transform.BasisX), transform.BasisZ)
    axis = DB.XYZ.BasisZ
    origin = DB.XYZ((box.Max.X + box.Min.X) / 2,
                    (box.Max.Y + box.Min.Y) / 2, 0.0)
    rotation_total = DB.Transform.CreateRotationAtPoint(axis, face_angle - box_angle, origin)
    box.Transform = box.Transform.Multiply(rotation_total)
    with revit.Transaction('BIM: Выравнивание границы 3D вида'):
        view.SetSectionBox(box)
        revit.uidoc.RefreshActiveView()


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    if isinstance(cur_view, DB.View3D) and cur_view.IsSectionBoxActive:
        orient_section_box(cur_view)
    elif isinstance(cur_view, DB.View3D) and not cur_view.IsSectionBoxActive:
        forms.alert("Границы 3D вида не включены.", exitscript=True)
    else:
        forms.alert('Должен быть открыт 3D вид.', exitscript=True)


script_execute()
