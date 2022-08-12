# coding=utf-8

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

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
        dependent_views = []
        for view in self.AllViews.Items:
            if view.IsChecked:
                views_to_copy.append(view.ViewToCheck)
        if views_to_copy:
            selected_view = self.DependentViews.SelectedItem.ViewToShow
            if selected_view:
                dependent_views = selected_view.GetDependentViewIds()

            with Transaction(doc, "BIM: Создать зависимые виды") as t:
                t.Start()
                for view in views_to_copy:
                    if selected_view:
                        for dep_view_id in dependent_views:
                            new_view_id = view.Duplicate(ViewDuplicateOption.AsDependent)

                            dep_view = doc.GetElement(dep_view_id)
                            template = dep_view.ViewTemplateId
                            cropbox = dep_view.CropBox
                            cropbox_vis = dep_view.CropBoxVisible

                            new_view = doc.GetElement(new_view_id)
                            new_view.ViewTemplateId = template
                            new_view.CropBox = cropbox
                            new_view.CropBoxVisible = cropbox_vis
                    else:
                        view.Duplicate(ViewDuplicateOption.AsDependent)
                t.Commit()
            self.DialogResult = True
            self.Close()
        else:
            alert("Не выбраны виды", exitscript=False)


class MainViewWithDependent():
    def __init__(self, view, name):
        self.ViewToShow = view
        self.Name = name


class MainView():
    def __init__(self, view, name):
        self.ViewToCheck = view
        self.Name = name
        self.IsChecked = False


def get_views(doc):
    all_views = FilteredElementCollector(doc).OfClass(ViewPlan).ToElements()
    # Main (not dependent) views with dependent views
    views_with_dependent = []
    # All main (not dependent) views
    main_views = []
    for view in all_views:
        if not view.IsTemplate:
            if view.CanViewBeDuplicated(ViewDuplicateOption.AsDependent):
                if view.GetPrimaryViewId() == ElementId.InvalidElementId:
                    view_check_box = MainView(view, view.Name)
                    main_views.append(view_check_box)
                    if view.GetDependentViewIds():
                        view_combo_box = MainViewWithDependent(view, view.Name)
                        views_with_dependent.append(view_combo_box)

    first_item = MainViewWithDependent(None, "<Без вида>")
    views_with_dependent.insert(0, first_item)

    main_views.sort(key=lambda view: view.Name)

    return views_with_dependent, main_views


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    dependent_views, main_views = get_views(doc)
    main_window = MainWindow(dependent_views, main_views)
    if not main_window.show_dialog():
        script.exit()


script_execute()