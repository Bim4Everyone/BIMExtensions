# -*- coding: utf-8 -*-

from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import script
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


def is_placeable(view):
    if view and view.ViewType and view.ViewType in [DB.ViewType.Schedule,
                                                    DB.ViewType.DraftingView,
                                                    DB.ViewType.Legend,
                                                    DB.ViewType.CostReport,
                                                    DB.ViewType.LoadsReport,
                                                    DB.ViewType.ColumnSchedule,
                                                    DB.ViewType.PanelSchedule]:
        return True
    return False


def update_if_placed(vport, exst_vps):
    for exst_vp in exst_vps:
        nameParam = exst_vp.Name
        if nameParam:
            if "жук" in nameParam.lower():
                exst_vp.SetBoxCenter(vport.GetBoxCenter())
                return True

    return False


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    allSheetedSchedules = DB.FilteredElementCollector(revit.doc) \
        .OfClass(DB.ScheduleSheetInstance) \
        .ToElements()

    selSheets = forms.select_sheets(title='Выберите нужные листы',
                                    button_name='Выбрать листы')

    # get a list of viewports to be copied, updated
    if selSheets and len(selSheets) > 0:
        if int(__revit__.Application.VersionNumber) > 2014:  # noqa
            cursheet = revit.uidoc.ActiveGraphicalView
            for v in selSheets:
                if cursheet.Id == v.Id:
                    selSheets.remove(v)
        else:
            cursheet = selSheets[0]
            selSheets.remove(cursheet)

        revit.uidoc.ActiveView = cursheet
        selected_vps = revit.pick_elements()

        if selected_vps:
            with revit.Transaction('Copy Viewports to Sheets'):
                for sht in selSheets:
                    existing_vps = [revit.doc.GetElement(x)
                                    for x in sht.GetAllViewports()]
                    existing_schedules = [x for x in allSheetedSchedules
                                          if x.OwnerViewId == sht.Id]
                    for vp in selected_vps:
                        nameParam = vp.Name
                        if nameParam:
                            if "жук" not in nameParam.lower():
                                continue

                        if isinstance(vp, DB.Viewport):
                            src_view = revit.doc.GetElement(vp.ViewId)

                            # check if viewport already exists
                            # and update location and type
                            update_if_placed(vp, existing_vps)

            show_executed_script_notification()
        else:
            forms.alert('Хотя бы одна легенда должна быть выбрана.')


script_execute()
