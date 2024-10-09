# -*- coding: utf-8 -*-
import clr
import datetime

from System.Collections.Generic import *

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

from System.Windows.Input import ICommand

import pyevent
from pyrevit import EXEC_PARAMS, revit
from pyrevit.forms import *
from pyrevit import script

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI.Selection import *

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep_libs.bim4everyone import *

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application


class CreateCommand(ICommand):
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
        # В Python при работе с событиями нужно явно
        # передавать импорт в обработчике события
        from System import EventArgs
        self._canExecuteChanged(self, EventArgs.Empty)

    def ViewModel_PropertyChanged(self, sender, e):
        self.OnCanExecuteChanged()

    def CanExecute(self, parameter):
        if self.__view_model.selected_start_level == 0:
            self.__view_model.error_text = "Выберите начальный уровень"
            return False

        if self.__view_model.selected_finish_level == 0:
            self.__view_model.error_text = "Выберите конечный уровень"
            return False

        if self.__view_model.selected_construction_stage == 0:
            self.__view_model.error_text = "Выберите стадию"
            return False

        self.__view_model.error_text = None
        return True

    def Execute(self, parameter):
        alert("Поехали!")
        return True



class SelectExcelFileCommand(ICommand):
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
        # В Python при работе с событиями нужно явно
        # передавать импорт в обработчике события
        from System import EventArgs
        self._canExecuteChanged(self, EventArgs.Empty)

    def ViewModel_PropertyChanged(self, sender, e):
        self.OnCanExecuteChanged()

    def CanExecute(self, parameter):
        return True

    def Execute(self, parameter):
        alert("Выбираем файл!")
        return True



class MainWindow(WPFWindow):
    def __init__(self):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), "MainWindow.xaml")
        super(MainWindow, self).__init__(self.xaml_source)

    def ButtonOK_Click(self, sender, e):
        self.DialogResult = True

    def ButtonCancel_Click(self, sender, e):
        self.DialogResult = False
        self.Close()


class MainWindowViewModel(Reactive):
    def __init__(self):
        Reactive.__init__(self)
        self.__levels = []
        self.__selected_start_level = 0
        self.__selected_finish_level = 0

        self.__construction_stages = []
        self.__selected_construction_stage = 0

        self.__error_text = ""

        self.__select_excel_file_command = SelectExcelFileCommand(self)
        self.__create_command = CreateCommand(self)

    @reactive
    def levels(self):
        return self.__levels

    @levels.setter
    def levels(self, value):
        self.__levels = value

    @reactive
    def selected_start_level(self):
        return self.__selected_start_level

    @selected_start_level.setter
    def selected_start_level(self, value):
        self.__selected_start_level = value


    @reactive
    def selected_finish_level(self):
        return self.__selected_finish_level

    @selected_finish_level.setter
    def selected_finish_level(self, value):
        self.__selected_finish_level = value


    @reactive
    def construction_stages(self):
        return self.__construction_stages

    @construction_stages.setter
    def construction_stages(self, value):
        self.__construction_stages = value


    @reactive
    def selected_construction_stage(self):
        return self.__selected_construction_stage

    @selected_construction_stage.setter
    def selected_construction_stage(self, value):
        self.__selected_construction_stage = value


    @reactive
    def error_text(self):
        return self.__error_text

    @error_text.setter
    def error_text(self, value):
        self.__error_text = value

    @property
    def create_command(self):
        return self.__create_command

    @property
    def select_excel_file_command(self):
        return self.__select_excel_file_command


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel()
    main_window.show_dialog()
    if not main_window.DialogResult:
        script.exit()


script_execute()