# -*- coding: utf-8 -*-

import clr
clr.AddReference('System')
clr.AddReference("System.Windows.Forms")

clr.AddReference('ClassLibrary2.dll')
clr.AddReference('MathNet.Numerics.dll')
clr.AddReference('Xceed.Wpf.Toolkit.dll')

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import System
from System.Windows.Input import ICommand
from Autodesk.Revit.DB import *

import pyevent  # pylint: disable=import-error
import os.path as op
from pyrevit import forms
from pyrevit import revit
from pyrevit import EXEC_PARAMS
from pyrevit.forms import Reactive, reactive

from dosymep_libs.bim4everyone import *


class MainWindow(forms.WPFWindow):
    def __init__(self, ):
        self._context = None
        self.xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')
        super(MainWindow, self).__init__(self.xaml_source)

    def button_ok_click(self, sender, e):
        self.Close()

    def button_cancel_click(self, sender, e):
        self.Close()


class MainWindowViewModel(Reactive):
    def __init__(self, *args):
        Reactive.__init__(self, *args)

        self.error_text = None
        self.__legend_name = None
        self.__legend_scale = "50"

        self.__create_legend_command = CreateLegendCommand(self)

    @reactive
    def error_text(self):
        return self.__error_text

    @error_text.setter
    def error_text(self, value):
        self.__error_text = value

    @reactive
    def legend_name(self):
        return self.__legend_name

    @legend_name.setter
    def legend_name(self, value):
        self.__legend_name = value

    @reactive
    def legend_scale(self):
        return self.__legend_scale

    @legend_scale.setter
    def legend_scale(self, value):
        self.__legend_scale = value

    @property
    def create_legend_command(self):
        return self.__create_legend_command


class CreateLegendCommand(ICommand):
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
        if not self.__view_model.legend_name:
            self.__view_model.error_text = "Заполните наименование легенды."
            return False

        if not self.__view_model.legend_scale:
            self.__view_model.error_text = "Заполните масштаб легенды."
            return False

        if not self.__is_int(self.__view_model.legend_scale):
            self.__view_model.error_text = "Масштаб легенды должен быть целым числом."
            return False

        if float(self.__view_model.legend_scale) <= 0:
            self.__view_model.error_text = "Масштаб легенды должен быть неотрицательным числом."
            return False

        self.__view_model.error_text = None
        return True

    def Execute(self, parameter):
        doc = __revit__.ActiveUIDocument.Document
        view = doc.ActiveView

        legends = [x for x in FilteredElementCollector(doc).OfClass(View) if x.ViewType == ViewType.Legend]
        legends_names = [x.Name for x in legends]
        legends = [x for x in legends if x.CanViewBeDuplicated(ViewDuplicateOption.Duplicate)]
        base_legend = legends[0]

        walls = [x for x in
                 FilteredElementCollector(doc, view.Id).OfClass(Wall).WhereElementIsNotElementType().ToElements()]

        scale = 1
        legend_name = self.__view_model.legend_name
        legend_scale = self.__view_model.legend_scale

        if not legend_name:
            legend_name = "Легенда по виду " + view.Name

        while legend_name in legends_names:
            legend_name = legend_name + " копия"

        transform = Transform.CreateTranslation(XYZ.Zero)
        transform = transform.ScaleBasis(scale)

        with revit.Transaction("BIM: Создание жука по плану этажа"):
            legend_id = base_legend.Duplicate(ViewDuplicateOption.Duplicate)
            legend = doc.GetElement(legend_id)

            legend.Name = legend_name
            legend.SetParamValue(BuiltInParameter.VIEW_SCALE_PULLDOWN_METRIC, legend_scale)

            for wall in walls:
                if not '(В)' in wall.Name:
                    curve = wall.Location.Curve
                    scaled_curve = curve.CreateTransformed(transform)

                    if scaled_curve:
                        doc.Create.NewDetailCurve(legend, scaled_curve)

    @staticmethod
    def __is_int(value):
        try:
            int(value)
            return True
        except:
            return False


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    main_window = MainWindow()
    main_window.DataContext = MainWindowViewModel()
    main_window.show_dialog()


script_execute()
