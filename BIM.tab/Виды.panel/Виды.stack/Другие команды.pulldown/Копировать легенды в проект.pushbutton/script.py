# -*- coding: utf-8 -*-

import os.path as op

from pyrevit import forms
from pyrevit import script
from pyrevit import EXEC_PARAMS
from pyrevit import revit, DB
from pyrevit.framework import List, Controls



from dosymep_libs.bim4everyone import *
from dosymep_libs.simple_services import *


doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application


class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


class doc_Option(object):
    def __init__(self, obj, state=False):
        self.state = state
        self.doc = obj
        self.name = obj.Title
        self.Title = obj.Title

    def __nonzero__(self):
        return self.state

    def __str__(self):
        return self.name


class SelectLevelFrom(forms.TemplateUserInputWindow):
    xaml_source = op.join(op.dirname(__file__), 'SelectFromCheckboxes.xaml')

    def _setup(self, **kwargs):
        self.checked_only = kwargs.get('checked_only', True)
        button_name = kwargs.get('button_name', None)
        if button_name:
            self.select_b.Content = button_name

        self.list_lb.SelectionMode = Controls.SelectionMode.Extended
        self.count_projects = kwargs.get('n_projects', 1)
        self.Height = 550
        if self.count_projects:
            self.Height = 250 + 30 * self.count_projects

        self._verify_context()
        self._list_options()

    def _verify_context(self):
        new_context = []
        for item in self._context:
            if not hasattr(item, 'state'):
                new_context.append(BaseCheckBoxItem(item))
            else:
                new_context.append(item)

        self._context = new_context

    def _list_options(self, checkbox_filter=None):
        if checkbox_filter:
            self.checkall_b.Content = 'Check'
            self.uncheckall_b.Content = 'Uncheck'
            self.toggleall_b.Content = 'Toggle'
            checkbox_filter = checkbox_filter.lower()
            self.list_lb.ItemsSource = \
                [checkbox for checkbox in self._context
                 if checkbox_filter in checkbox.name.lower()]
        else:
            self.checkall_b.Content = 'Выделить все'
            self.uncheckall_b.Content = 'Сбросить выделение'
            self.toggleall_b.Content = 'Инвертировать'
            self.list_lb.ItemsSource = self._context

    def _set_states(self, state=True, flip=False, selected=False):
        all_items = self.list_lb.ItemsSource
        if selected:
            current_list = self.list_lb.SelectedItems
        else:
            current_list = self.list_lb.ItemsSource
        for checkbox in current_list:
            if flip:
                checkbox.state = not checkbox.state
            else:
                checkbox.state = state

        # push list view to redraw
        self.list_lb.ItemsSource = None
        self.list_lb.ItemsSource = all_items

    def toggle_all(self, sender, args):
        """Handle toggle all button to toggle state of all check boxes."""
        self._set_states(flip=True)

    def check_all(self, sender, args):
        """Handle check all button to mark all check boxes as checked."""
        self._set_states(state=True)

    def uncheck_all(self, sender, args):
        """Handle uncheck all button to mark all check boxes as un-checked."""
        self._set_states(state=False)

    def check_selected(self, sender, args):
        """Mark selected checkboxes as checked."""
        self._set_states(state=True, selected=True)

    def uncheck_selected(self, sender, args):
        """Mark selected checkboxes as unchecked."""
        self._set_states(state=False, selected=True)

    def button_select(self, sender, args):
        """Handle select button click."""
        if self.checked_only:
            self.response = [x for x in self._context if x.state]
        else:
            self.response = self._context
        self.response = {'docs': self.response}
        self.Close()


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    open_docs = [d for d in revit.docs if not d.IsLinked]
    open_docs.remove(revit.doc)
    if len(open_docs) < 1:
        forms.alert('Должно быть открыто минимум два проекта.', exitscript=True)

    selection = [x for x in revit.get_selection()
                 if x.ViewType == DB.ViewType.Legend]

    if not selection:
        forms.alert("Должна быть выбрана как минимум одна легенда.", exitscript=True)

    o_docs = [doc_Option(doc) for doc in open_docs]
    res = SelectLevelFrom.show(o_docs,
                               n_projects=len(o_docs),
                               title='Выберите проекты',
                               button_name='Копировать легенды')

    if not res:
        script.exit()

    docs_2_process = [doc for doc in res['docs'] if doc.state]
    if not docs_2_process:
        script.exit()

    for dest_doc in docs_2_process:
        legends_project = []
        base_legend_view = None

        for v in DB.FilteredElementCollector(dest_doc.doc).OfClass(DB.View):
            if v.ViewType == DB.ViewType.Legend:
                legends_project.append(v.Name)
                base_legend_view = v

        if base_legend_view is None:
            forms.alert('В проекте "{0}" должна быть минимум одна легенда.'.format(dest_doc.Title), exitscript=True)

        for srcView in selection:
            viewElements = \
                DB.FilteredElementCollector(revit.doc, srcView.Id) \
                    .ToElements()

            element_list = []
            for el in viewElements:
                if isinstance(el, DB.Element) and el.Category:
                    element_list.append(el.Id)

            if len(element_list) < 1:
                print('Проверка содержимого {0}. Элементы в легенде не найдены.'.format(srcView.Title))
                continue

            with revit.Transaction('Copy Legends to this document', doc=dest_doc.doc):
                dest_view = dest_doc.doc.GetElement(base_legend_view.Duplicate(DB.ViewDuplicateOption.Duplicate))

                options = DB.CopyPasteOptions()
                options.SetDuplicateTypeNamesHandler(CopyUseDestination())
                copied_element = \
                    DB.ElementTransformUtils.CopyElements(
                        srcView,
                        List[DB.ElementId](element_list),
                        dest_view,
                        None,
                        options)

                for dest, src in zip(copied_element, element_list):
                    dest_view.SetElementOverrides(dest, srcView.GetElementOverrides(src))

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