# -*- coding: utf-8 -*-
import clr
clr.AddReference("OpenMcdf.dll")
clr.AddReference("dosymep.Revit.dll")

from dosymep.Revit import *
from pyrevit.forms import *
from pyrevit import EXEC_PARAMS

from Autodesk.Revit.DB import *

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument


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
            link.is_checked = not (link.is_checked)


class LinkedFile(Reactive):
    def __init__(self, revit_link):
        self.__link_name = revit_link.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString()
        self.__link_status = ""
        self.__is_checked = False

    @reactive
    def link_name(self):
        return self.__link_name

    @link_name.setter
    def link_name(self, value):
        self.__link_name = value

    @reactive
    def link_status(self):
        return self.__link_status

    @link_status.setter
    def link_status(self, value):
        self.__link_status = value

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

    return all_links



@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    links = get_links_from_document(doc)
    main_window = MainWindow(links)
    main_window.show_dialog()


script_execute()
