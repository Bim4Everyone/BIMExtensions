# -*- coding: utf-8 -*-

import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from Autodesk.Revit import DB

from pyrevit.framework import *
from pyrevit import forms
from pyrevit import script
from pyrevit import revit
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *


class Option(object):
    def __init__(self, obj, state=False):
        self.state = state
        self.name = obj.Name
        self.elevation = obj.Elevation

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

        self.Height = 550
        # for i in range(1,4):
        #	self.purpose.AddText(str(i))

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
        self.response = {'level': self.response}
        self.Close()


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    lvls = DB.FilteredElementCollector(revit.doc).OfClass(DB.Level)
    ops = [Option(x) for x in lvls]
    ops.sort(key=lambda x: x.elevation)
    res = SelectLevelFrom.show(ops,
                               button_name='ОК', title="Выберите уровни")

    LEVEL = []
    if res:
        LEVEL = [x.name for x in res['level']]
    matchlist = []

    selection = revit.get_selection()
    # print LEVEL
    if len(selection) > 0 and len(LEVEL) > 0:
        # print selection
        for el in selection:
            # print el
            if isinstance(el, DB.Wall):

                family = el.WallType.FamilyName
                wallSet = [x for x in DB.FilteredElementCollector(revit.doc).OfClass(DB.Wall).ToElements() if
                           x.WallType.FamilyName == family and revit.doc.GetElement(
                               x.LevelId).Name in LEVEL and x.GetParamValueOrDefault(
                               DB.BuiltInParameter.WALL_BASE_CONSTRAINT)]

                # break
                for wall in wallSet:
                    matchlist.append(wall.Id)
            else:
                if hasattr(el, "Symbol"):
                    family = el.Symbol.Family
                    symbolIdSet = family.GetFamilySymbolIds()

                    for symid in symbolIdSet:
                        cl = DB.FilteredElementCollector(revit.doc) \
                            .WherePasses(DB.FamilyInstanceFilter(revit.doc, symid)) \
                            .ToElements()

                        for el in cl:
                            level = revit.doc.GetElement(el.LevelId)
                            host = el.Host

                            if level:
                                if level.Name in LEVEL:
                                    matchlist.append(el.Id)
                            elif host:
                                host
                                if host:
                                    level = revit.doc.GetElement(host.LevelId)
                                    if level:
                                        if level.Name in LEVEL:
                                            matchlist.append(el.Id)
                            else:
                                matchlist = matchlist + list(cl)
                                break

        selection.set_to(matchlist)


script_execute()
