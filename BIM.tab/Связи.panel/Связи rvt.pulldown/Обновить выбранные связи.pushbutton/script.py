# -*- coding: utf-8 -*-
import clr
clr.AddReference("OpenMcdf.dll")
clr.AddReference("dosymep.Revit.dll")

from dosymep.Revit import *
from pyrevit.forms import *
from pyrevit import EXEC_PARAMS

import pyevent
from System.Windows.Input import ICommand

from Autodesk.Revit.DB import *

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


class PickFolderCommand(ICommand):
    CanExecuteChanged, _canExecuteChanged = pyevent.make_event()

    def __init__(self, view_model, *args):
        ICommand.__init__(self, *args)
        self.__view_model = view_model

    def add_CanExecuteChanged(self, value):
        self.CanExecuteChanged += value

    def remove_CanExecuteChanged(self, value):
        self.CanExecuteChanged -= value

    def OnCanExecuteChanged(self):
        self._canExecuteChanged(self, System.EventArgs.Empty)

    def CanExecute(self, parameter):
        return True

    def Execute(self, parameter):
        picked_folder = pick_folder()
        if picked_folder:
            self.__view_model.folder_path = picked_folder


class UpdateLinks(ICommand):
    pass


class MainWindowViewModel(Reactive):
    def __init__(self, *args):
        Reactive.__init__(self, *args)

        self.__folder_path = ""
        self.__pick_folder_command = PickFolderCommand(self)

    @property
    def PickFolderCommand(self):
        return self.__pick_folder_command

    @reactive
    def folder_path(self):
        return self.__folder_path

    @folder_path.setter
    def folder_path(self, value):
        self.__folder_path = value


class MainWindow(WPFWindow):
    def __init__(self, links):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')
        super(MainWindow, self).__init__(self.xaml_source)

        self.revit_links.ItemsSource = links

    def update_states(self, value):
        for link in self.revit_links.ItemsSource:
            link.is_checked = value

    def select_all(self, sender, args):
        self.update_states(True)

    def deselect_all(self, sender, args):
        self.update_states(False)

    def invert(self, sender, args):
        for link in self.revit_links.ItemsSource:
            link.is_checked = not link.is_checked


class LinkedFile(Reactive):
    def __init__(self, revit_link):
        self.link = revit_link
        self.link_name = revit_link.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString()

        status = revit_link.GetLinkedFileStatus()
        self.link_status = status
        if status == LinkedFileStatus.NotFound or status == LinkedFileStatus.Unloaded:
            self.is_checked = True
        else:
            self.is_checked = False

    @reactive
    def link_status(self):
        return self.__link_status

    @link_status.setter
    def link_status(self, value):
        if value == LinkedFileStatus.Loaded:
            self.__link_status = "Загружена"
        elif value == LinkedFileStatus.Unloaded:
            self.__link_status = "Не загружена"
        elif value == LinkedFileStatus.NotFound:
            self.__link_status = "Не найдена"
        elif value == LinkedFileStatus.LocallyUnloaded:
            self.__link_status = "Выгружена локально"
        else:
            self.__link_status = "???"

    @reactive
    def is_checked(self):
        return self.__is_checked

    @is_checked.setter
    def is_checked(self, value):
        self.__is_checked = value


def get_links_from_document(document):
    links = FilteredElementCollector(document).OfClass(RevitLinkType).ToElements()
    all_links = []
    for link in links:
        if not link.IsNestedLink:
            linked_file = LinkedFile(link)
            all_links.append(linked_file)
    # add links sorting
    return all_links


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    links = get_links_from_document(doc)
    main_window = MainWindow(links)
    main_window.DataContext = MainWindowViewModel()
    main_window.show_dialog()


script_execute()
