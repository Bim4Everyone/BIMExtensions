# coding=utf-8

from pyrevit import revit, DB, EXEC_PARAMS
from pyrevit import script

from dosymep_libs.bim4everyone import log_plugin
from dosymep_libs.simple_services import notification


output = script.get_output()

# collect target views
selection = revit.get_selection()
target_views = selection.elements
target_views = [view for view in target_views if isinstance(view, DB.View)]
if not target_views:
    target_views = [revit.active_view]


def shake_dimensions(target_view):
    dimensions = \
        DB.FilteredElementCollector(target_view.Document, target_view.Id)\
          .OfClass(DB.Dimension)\
          .WhereElementIsNotElementType()\
          .ToElements()

    print("Встряхнуть размеры на виде: {}"
          .format(revit.query.get_name(target_view)))

    for dimension in dimensions:
        with revit.Transaction("BIM: Встряхнуть размеры"):
            pinned = dimension.Pinned

            dimension.Pinned = False
            dimension.Location.Move(DB.XYZ(0.1, 0.1, 0.1))
            dimension.Location.Move(DB.XYZ(-0.1, -0.1, -0.1))

            dimension.Pinned = pinned


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    print("Обновление размеров на {} видах".format(len(target_views)))
    with revit.TransactionGroup("BIM: Встряхнуть размеры"):
        for idx, view in enumerate(target_views):
            shake_dimensions(view)
            output.update_progress(idx+1, len(target_views))

    print("Все размеры были обновлены ...")


script_execute()
