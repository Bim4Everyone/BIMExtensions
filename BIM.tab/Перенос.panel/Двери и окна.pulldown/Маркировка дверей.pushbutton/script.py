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

param_name_for_filter = "ФОП_Учитывать в спецификации"

doors_name_start_with = "Двр_Двр"
hatches_name_start_with = "Двр_Люк"

param_name_for_mark = "ФОП_Марка типа"

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

        if self.__view_model.selected_construction_phase == 0:
            self.__view_model.error_text = "Выберите стадию"
            return False

        self.__view_model.error_text = None
        return True

    def Execute(self, parameter):

        doors = (FilteredElementCollector(doc)
                 .OfCategory(BuiltInCategory.OST_Doors)
                 .WhereElementIsNotElementType()
                 .ToElements())

        start_level_elevation = self.__view_model.selected_start_level.Elevation
        finish_level_elevation = self.__view_model.selected_finish_level.Elevation
        phase_id = self.__view_model.selected_construction_phase.Id

        print("start_level_elevation")
        print(start_level_elevation)
        print("finish_level_elevation")
        print(finish_level_elevation)
        print("phase_id")
        print(phase_id)

        door_types_for_work = []
        hatch_types_for_work = []
        for door in doors:
            print(door)
            level = doc.GetElement(door.LevelId)
            level_elevation = level.Elevation

            print(level_elevation)

            # Фильтрация по уровню
            if finish_level_elevation >= level_elevation >= start_level_elevation:
                door_phase_id = door.CreatedPhaseId
                # Фильтрация по стадии
                if door_phase_id == phase_id:
                    door_type = doc.GetElement(door.GetTypeId())
                    # Фильтрация по параметру "ФОП_Учитывать в спецификации"
                    if door_type.GetParamValue(param_name_for_filter):
                        door_family_name = door_type.FamilyName
                        # Фильтрация по имени семейства
                        if doors_name_start_with in door_family_name:
                            door_types_for_work.append(door_type)
                        elif hatches_name_start_with in door_family_name:
                            hatch_types_for_work.append(door_type)

        self.clear_mark(door_types_for_work)
        self.clear_mark(hatch_types_for_work)

        door_types_for_work = self.get_unique_items_by_id(door_types_for_work)
        hatch_types_for_work = self.get_unique_items_by_id(hatch_types_for_work)

        print("door_types_for_work")
        print(door_types_for_work)
        print("hatch_types_for_work")
        print(hatch_types_for_work)
        return True

    def clear_mark(self, elements):
        with revit.Transaction("BIM: Очистка марок"):
            for element in elements:
                element.SetParamValue(param_name_for_mark, "")

    def get_unique_items_by_id(self, elems):
        d = {}
        for item in elems:
            id_as_str = str(item.Id)
            d[id_as_str] = item
        return d.values()

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

        self.__construction_phases = []
        self.__selected_construction_phase = 0

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
    def construction_phases(self):
        return self.__construction_phases

    @construction_phases.setter
    def construction_phases(self, value):
        self.__construction_phases = value


    @reactive
    def selected_construction_phase(self):
        return self.__selected_construction_phase

    @selected_construction_phase.setter
    def selected_construction_phase(self, value):
        self.__selected_construction_phase = value


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

    def get_levels(self):
        levels_in_pj = (FilteredElementCollector(doc)
                                         .OfCategory(BuiltInCategory.OST_Levels)
                                         .WhereElementIsNotElementType()
                                         .ToElements())
        self.levels = levels_in_pj
        if len(levels_in_pj) == 0:
            alert("Не найдено ни одного уровня!")
        self.selected_start_level = levels_in_pj[0]
        self.selected_finish_level = levels_in_pj[len(levels_in_pj) - 1]

    def get_phases(self):
        phases = doc.Phases
        self.construction_phases = phases
        self.selected_construction_phase = phases[0]

@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    main_window = MainWindow()

    main_window_view_model = MainWindowViewModel()
    main_window_view_model.get_levels()
    main_window_view_model.get_phases()

    main_window.DataContext = main_window_view_model

    main_window.show_dialog()
    if not main_window.DialogResult:
        script.exit()


script_execute()