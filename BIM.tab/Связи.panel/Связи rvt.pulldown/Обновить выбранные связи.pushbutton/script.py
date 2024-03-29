# -*- coding: utf-8 -*-
import clr

clr.AddReference("OpenMcdf.dll")
clr.AddReference("dosymep.Revit.dll")

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


class UpdateLinksCommand(ICommand):
    CanExecuteChanged, _canExecuteChanged = pyevent.make_event()

    def __init__(self, view_model, *args):
        ICommand.__init__(self, *args)
        self.__view_model = view_model
        self.__view_model.PropertyChanged += self.ViewModel_PropertyChanged

    def add_CanExecuteChanged(self, value):
        self.CanExecuteChanged += value

    def remove_CanExecuteChanged(self, value):
        self.CanExecuteChanged -= value

    def OnCanExecuteChanged(self):
        self._canExecuteChanged(self, System.EventArgs.Empty)

    def ViewModel_PropertyChanged(self, sender, e):
        self.OnCanExecuteChanged()

    def CanExecute(self, parameter):
        if not self.__view_model.folder_path:
            self.__view_model.error_text = "Необходимо выбрать путь к папке."
            return False

        links_to_update = [x for x in self.__view_model.links if x.is_checked]
        if not links_to_update:
            self.__view_model.error_text = "Необходимо выбрать связи."
            return False

        self.__view_model.error_text = None
        return True

    def __find_file(self, main_path, link_name, level=0):
        if level < 5:
            files = os.listdir(main_path)
            if link_name in files:
                return os.path.join(main_path, link_name)
            else:
                level += 1
                for file_in_dir in files:
                    sub_path = os.path.join(main_path, file_in_dir)
                    if os.path.isdir(sub_path):
                        result = self.__find_file(sub_path, link_name, level)
                        if result:
                            return result

    def __filter_links(self, links):
        filtered_links = []
        for link in links:
            if link.is_checked:
                if link.revit_link.GetLinkedFileStatus() != LinkedFileStatus.Loaded:
                    if link.revit_link.IsFromLocalPath():
                        if link.revit_link.IsNotLoadedIntoMultipleOpenDocuments():
                            filtered_links.append(link)
        return filtered_links

    def Execute(self, parameter):
        ws_config = WorksetConfiguration()
        links_to_update = self.__filter_links(self.__view_model.links)
        for link in links_to_update:
            link_path = self.__find_file(self.__view_model.folder_path, link.link_name)
            if link_path:
                revit_path = FilePath(link_path)
                link.revit_link.LoadFrom(revit_path, ws_config)


class InvertCommand(ICommand):
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
        for link in self.__view_model.links:
            if link.is_ws_open:
                link.is_checked = not link.is_checked


class UpdateStatesCommand(ICommand):
    CanExecuteChanged, _canExecuteChanged = pyevent.make_event()

    def __init__(self, view_model, value, *args):
        ICommand.__init__(self, *args)
        self.__view_model = view_model
        self.__value = value

    def add_CanExecuteChanged(self, value):
        self.CanExecuteChanged += value

    def remove_CanExecuteChanged(self, value):
        self.CanExecuteChanged -= value

    def OnCanExecuteChanged(self):
        self._canExecuteChanged(self, System.EventArgs.Empty)

    def CanExecute(self, parameter):
        return True

    def Execute(self, parameter):
        for link in self.__view_model.links:
            if link.is_ws_open:
                link.is_checked = self.__value


class MainWindowViewModel(Reactive):
    def __init__(self, links):
        Reactive.__init__(self)

        self.__links = links
        self.__pick_folder_command = PickFolderCommand(self)
        self.__update_links_command = UpdateLinksCommand(self)
        self.__invert_command = InvertCommand(self)
        self.__set_true_command = UpdateStatesCommand(self, True)
        self.__set_false_command = UpdateStatesCommand(self, False)
        self.__error_text = ""

    @property
    def PickFolderCommand(self):
        return self.__pick_folder_command

    @property
    def UpdateLinksCommand(self):
        return self.__update_links_command

    @property
    def InvertCommand(self):
        return self.__invert_command

    @property
    def SetTrueCommand(self):
        return self.__set_true_command

    @property
    def SetFalseCommand(self):
        return self.__set_false_command

    @reactive
    def folder_path(self):
        return self.__folder_path

    @folder_path.setter
    def folder_path(self, value):
        self.__folder_path = value

    @reactive
    def links(self):
        return self.__links

    @links.setter
    def links(self, value):
        self.__links = value

    @reactive
    def error_text(self):
        return self.__error_text

    @error_text.setter
    def error_text(self, value):
        self.__error_text = value


class MainWindow(WPFWindow):
    def __init__(self):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')
        super(MainWindow, self).__init__(self.xaml_source)

    def ButtonOK_Click(self, sender, e):
        self.DialogResult = True
        self.Close()

    def ButtonCancel_Click(self, sender, e):
        self.DialogResult = False
        self.Close()


class LinkedFile(Reactive):
    def __init__(self, revit_link):
        self.revit_link = revit_link
        self.link_name = revit_link.Parameter[BuiltInParameter.ALL_MODEL_TYPE_NAME].AsString()

        status = revit_link.GetLinkedFileStatus()
        self.link_status = status

        if status == LinkedFileStatus.NotFound or status == LinkedFileStatus.Unloaded:
            self.is_checked = True
        else:
            self.is_checked = False

        document = revit_link.Document
        workset_table = document.GetWorksetTable()
        workset_id = document.GetWorksetId(self.revit_link.Id)
        workset = workset_table.GetWorkset(workset_id)

        self.is_ws_open = workset.IsOpen
        self.ws_status = workset.IsOpen


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
    def ws_status(self):
        return self.__ws_status

    @ws_status.setter
    def ws_status(self, value):
        if value:
            self.__ws_status = "Открыт"
        else:
            self.__ws_status = "Закрыт"

    @reactive
    def is_ws_open(self):
        return self.__is_ws_open

    @is_ws_open.setter
    def is_ws_open(self, value):
        self.__is_ws_open = value

    @reactive
    def is_checked(self):
        return self.__is_checked

    @is_checked.setter
    def is_checked(self, value):
        self.__is_checked = value


def get_links_from_document(document):
    all_links = FilteredElementCollector(document).OfClass(RevitLinkType).ToElements()
    links = [LinkedFile(x) for x in all_links if not x.IsNestedLink]
    links = sorted(links, key=lambda x: x.link_name)
    return links


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    links = get_links_from_document(doc)
    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel(links)
    if not main_window.show_dialog():
        script.exit()


script_execute()
