# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import os.path as op

from Autodesk.Revit.DB import *

from pyrevit import forms
from pyrevit import script
from pyrevit import revit, DB
from pyrevit import EXEC_PARAMS
from pyrevit.framework import List

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

project_params = ProjectParameters.Create(app)
project_params.SetupRevitParams(doc, ProjectParamsConfig.Instance.ViewGroup)


class SelectLevelFrom(forms.TemplateUserInputWindow):
    xaml_source = op.join(op.dirname(__file__), 'SelectFromComboBox.xaml')

    def _setup(self, **kwargs):
        button_name = kwargs.get('button_name', None)
        self.views = kwargs.get('views', None)
        if button_name:
            self.select_b.Content = button_name

        for i in kwargs['views']:
            self.purpose.AddText(i)

    def button_select(self, sender, args):
        """Handle select button click."""
        if self.purpose.Text:
            purpose = self.purpose.Text
        else:
            purpose = ""

        self.response = {"purpose": purpose}
        self.Close()


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selection = [x for x in revit.get_selection()
                 if x.ViewType == DB.ViewType.Legend]

    if not selection:
        forms.alert("Должна быть выбрана минимум одна легенда.", exitscript=True)

    poject_drafts = []
    doc_views = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Views).ToElements()

    views = []
    for view in doc_views:
        view_group = view.GetParamValueOrDefault(ProjectParamsConfig.Instance.ViewGroup)
        if view_group and view_group not in views:
            views.append(view_group)

    response = SelectLevelFrom.show([],
                                    title="Выберите назначение вида",
                                    button_name="Ок",
                                    views=views)

    if not response:
        script.exit()

    drafting_view = False
    for v in DB.FilteredElementCollector(revit.doc).OfClass(DB.View):
        if v.ViewType == DB.ViewType.DraftingView:
            drafting_view = True
            draftingViewType = v
            poject_drafts.append(v.Name)

    if not drafting_view:
        forms.alert('В проекте должен быть минимум один чертежный вид.', exitscript=True)

    for src_view in selection:
        view_elements = DB.FilteredElementCollector(revit.doc, src_view.Id) \
            .ToElements()

        element_list = []
        for el in view_elements:
            if isinstance(el, DB.Element) \
                    and el.Category \
                    and el.Category.Id != ElementId(BuiltInCategory.OST_LegendComponents):
                element_list.append(el.Id)

        if len(element_list) < 1:
            print('Пропуск {0}. Никаких элементов не найдено.'.format(src_view.Name))
            continue

        with revit.Transaction('Duplicate Legend as Drafting'):
            dest_view = revit.doc.GetElement(draftingViewType.Duplicate(DB.ViewDuplicateOption.Duplicate))

            options = DB.CopyPasteOptions()
            options.SetDuplicateTypeNamesHandler(CopyUseDestination())

            copiedElement = \
                DB.ElementTransformUtils.CopyElements(
                    src_view,
                    List[DB.ElementId](element_list),
                    dest_view,
                    None,
                    options)

            for dest, src in zip(copiedElement, element_list):
                dest_view.SetElementOverrides(dest, src_view.GetElementOverrides(src))

            if src_view.Name not in poject_drafts:
                dest_view.Name = src_view.Name
                poject_drafts.append(dest_view.Name)
            else:
                index = 1
                while '{}-{}'.format(src_view.Name, index) in poject_drafts:
                    index += 1
                dest_view.Name = '{}-{}'.format(src_view.Name, index)
                poject_drafts.append(dest_view.Name)

            if response:
                purpose = response["purpose"]
                dest_view.Scale = src_view.Scale
                dest_view.SetParamValue(ProjectParamsConfig.Instance.ViewGroup, purpose)

    show_executed_script_notification()


script_execute()