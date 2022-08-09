# coding=utf-8

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import pyevent #pylint: disable=import-error
from System.Windows.Input import ICommand

from dosymep_libs.bim4everyone import *

from pyrevit import *
from pyrevit.forms import *
from pyrevit.revit import *

from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document
uiapp = __revit__
app = uiapp.Application


class MainWindow(WPFWindow):
    def __init__(self, dependent_views, main_views):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')
        super(MainWindow, self).__init__(self.xaml_source)

        self.DependentViews.ItemsSource = dependent_views
        self.DependentViews.DisplayMemberPath = "Name"
        self.DependentViews.SelectedIndex = 0
        self.AllViews.ItemsSource = main_views

    def create_dependent_views(self, sender, args):
        views_to_copy = []
        for view in self.AllViews.Items:
            if view.IsChecked:
                views_to_copy.append(view.ViewToCheck)
        if views_to_copy:
            selected_view = self.DependentViews.SelectedItem.ViewToShow
            if selected_view:
                template = selected_view.ViewTemplateId
                cropbox = selected_view.CropBox

            with Transaction(doc, "Name") as t:
                t.Start()
                for view in views_to_copy:
                    if view.GetPrimaryViewId().IntegerValue == -1 and len(view.GetDependentViewIds()) == 0:
                        new_view = view.Duplicate(ViewDuplicateOption.AsDependent)
                        if selected_view:
                            doc.GetElement(new_view).ViewTemplateId = template
                            doc.GetElement(new_view).CropBox = cropbox
                            doc.GetElement(new_view).CropBoxVisible = False
                t.Commit()
            self.DialogResult = True
            self.Close()
        else:
            alert("Не выбраны виды", exitscript=False)


class ViewInComboBox():
    def __init__(self, view, name):
        self.ViewToShow = view
        self.Name = name


class ViewInCheckBox():
    def __init__(self, view, name):
        self.ViewToCheck = view
        self.Name = name
        self.IsChecked = False


def GetViews(doc):
    all_views = FilteredElementCollector(doc).OfClass(ViewPlan).ToElements()
    dependent_views = []
    main_views = []
    for view in all_views:
        if view.CanViewBeDuplicated(ViewDuplicateOption.AsDependent) and not view.IsTemplate:
            if view.GetPrimaryViewId() == ElementId.InvalidElementId:
                view_check_box = ViewInCheckBox(view, view.Name)
                main_views.append(view_check_box)
            else:
                primary_view_id = view.GetPrimaryViewId()
                primary_view = doc.GetElement(primary_view_id)
                full_name = "{}: {}".format(view.Name, primary_view.Name)
                view_combo_box = ViewInComboBox(view, full_name)
                dependent_views.append(view_combo_box)

    first_item = ViewInComboBox(None, "<Не копировать настройки>")
    dependent_views.insert(0, first_item)

    main_views.sort(key=lambda view: view.Name)

    return dependent_views, main_views


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    dependent_views, main_views = GetViews(doc)
    main_window = MainWindow(dependent_views, main_views)
    if not main_window.show_dialog():
        script.exit()


script_execute()