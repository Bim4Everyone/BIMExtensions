# -*- coding: utf-8 -*-

from pyrevit import forms
from pyrevit import script
from pyrevit import revit, DB
from pyrevit import EXEC_PARAMS
from pyrevit.framework import List

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

logger = script.get_logger()


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    # get a list of selected drafting views
    selection = [el for el in revit.get_selection()
                 if isinstance(el, DB.View)
                 and el.ViewType == DB.ViewType.DraftingView]

    if not selection:
        forms.alert('Должен быть выбран как минимум один чертежный вид.', exitscript=True)

    legends_project = []
    base_legend_view = None
    for v in DB.FilteredElementCollector(revit.doc).OfClass(DB.View):
        if v.ViewType == DB.ViewType.Legend:
            base_legend_view = v
            legends_project.append(v.Name)

    if not base_legend_view:
        forms.alert('В модели должна быть как минимум одна легенда.', exitscript=True)

    for srcView in selection:
        view_elements = \
            DB.FilteredElementCollector(revit.doc, srcView.Id).ToElements()

        elements_to_copy = []
        for el in view_elements:
            if isinstance(el, DB.Element) and el.Category:
                elements_to_copy.append(el.Id)

        if len(elements_to_copy) < 1:
            logger.debug('Пропуск {0}. Никаких копируемых элементов не найдено.'.format(srcView.ViewName))
            continue

        with revit.Transaction('Duplicate Drafting as Legend'):
            dest_view = revit.doc.GetElement(base_legend_view.Duplicate(DB.ViewDuplicateOption.Duplicate))

            options = DB.CopyPasteOptions()
            options.SetDuplicateTypeNamesHandler(CopyUseDestination())
            copied_element = \
                DB.ElementTransformUtils.CopyElements(
                    srcView,
                    List[DB.ElementId](elements_to_copy),
                    dest_view,
                    None,
                    options)

            for dest, src in zip(copied_element, elements_to_copy):
                dest_view.SetElementOverrides(dest,
                                              srcView.GetElementOverrides(src))

            if srcView.Name not in legends_project:
                dest_view.Name = srcView.Name
                legends_project.append(dest_view.Name)
            else:
                index = 1
                while '{}-{}'.format(srcView.Name, index) in legends_project:
                    index += 1
                dest_view.Name = '{}-{}'.format(srcView.Name, index)
                legends_project.append(dest_view.Name)

            dest_view.Scale = srcView.Scale

    show_executed_script_notification()


script_execute()