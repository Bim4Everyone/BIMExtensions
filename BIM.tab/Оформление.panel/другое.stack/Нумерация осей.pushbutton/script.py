# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
import clr
import math

clr.AddReference('System')
clr.AddReference("System.Windows.Forms")

from System.Windows.Forms import MessageBox
from System.Collections.Generic import List
from Autodesk.Revit.DB import *
from Autodesk.Revit.Creation import ItemFactoryBase

from pyrevit import forms
from pyrevit import revit
from pyrevit import EXEC_PARAMS
from pyrevit.framework import Controls

from dosymep_libs.bim4everyone import *


class SelectLevelFrom(forms.TemplateUserInputWindow):
    xaml_source = op.join(op.dirname(__file__), 'MainWindow.xaml')

    def _setup(self, **kwargs):
        self.checked_only = kwargs.get('checked_only', True)
        button_name = kwargs.get('button_name', None)
        self.hide_element(self.clrsuffix_b)
        self.hide_element(self.clrprefix_b)
        self.hide_element(self.clrstart_b)
        if button_name:
            self.select_b.Content = button_name

    def suffix_txt_changed(self, sender, args):
        """Handle text change in search box."""
        if self.suffix.Text == '':
            self.hide_element(self.clrsuffix_b)
        else:
            self.show_element(self.clrsuffix_b)

    def clear_suffix(self, sender, args):
        """Clear search box."""
        self.suffix.Text = ''
        self.suffix.Clear()
        self.suffix.Focus

    def prefix_txt_changed(self, sender, args):
        """Handle text change in search box."""
        if self.prefix.Text == '':
            self.hide_element(self.clrprefix_b)
        else:
            self.show_element(self.clrprefix_b)

    def clear_prefix(self, sender, args):
        """Clear search box."""
        self.prefix.Text = ''
        self.prefix.Clear()
        self.prefix.Focus

    def start_txt_changed(self, sender, args):
        """Handle text change in search box."""
        if self.suffix.Text == '':
            self.hide_element(self.clrstart_b)
        else:
            self.show_element(self.clrstart_b)

    def clear_start(self, sender, args):
        """Clear search box."""
        self.start.Text = ''
        self.start.Clear()
        self.start.Focus

    def button_select(self, sender, args):
        """Handle select button click."""

        self.response = {'start': self.start.Text,
                         'prefix': self.prefix.Text,
                         'suffix': self.suffix.Text,
                         'isReverse': self.isReverse.IsChecked}
        self.Close()


class GridElement:
    def __init__(self, grid, isDigit=True):
        self.__grid = grid
        self.direction = grid.Curve.Direction
        self.origin = grid.Curve.Origin
        self.isDigit = isDigit
        self.endPoint = grid.Curve.GetEndPoint
        # self.angle0 = self.direction.AngleTo(XYZ(1,0,0))
        if self.endPoint(1).X - self.endPoint(0).X == 0:
            self.angle = 90
        else:
            self.angle = round(
                (math.atan((self.endPoint(1).Y - self.endPoint(0).Y) / (self.endPoint(1).X - self.endPoint(0).X))) * (
                        180 / (math.pi)), 3)

        self.name = self.__grid.Name

    def getRangeY(self):
        direction = self.direction
        origin = self.origin

        res = (0 - origin.X) * (direction.Y / direction.X) + origin.Y

        return res

    def getRangeX(self):
        direction = self.direction
        origin = self.origin
        res = (0 - origin.Y) * (direction.X / direction.Y) + origin.X
        # print res
        # print '-------------------------------------------------------------------'

        return res

    def getRange(self):
        x = self.getRangeX()
        y = self.getRangeY()
        # print 'Direction x: {} - y:{}'.format(self.direction.X,self.direction.Y)
        res = math.sqrt(x * x + y * y)

        if self.isDigit:
            if x < 0:
                res = -res
        elif y < 0:
            res = -res

        print res
        return res

    def setName(self, name):
        self.__grid.Name = name
        self.name = name


class GridElementContainer:
    def __init__(self, gridElementList):
        self.__gridElementList = gridElementList


def addChar(lst, CHARACTERS):
    if lst:
        res = lst[::]
        if res[-1] + 1 <= len(CHARACTERS):
            res[-1] += 1
            return res
        else:
            return addChar(lst[:-1], CHARACTERS) + [1]
    else:
        return [1]


def flipList(list):
    n = len(list)
    result = [None] * n
    for i in range(n):
        result[i] = list[n - 1 - i]
    return result


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    doc = __revit__.ActiveUIDocument.Document
    uidoc = __revit__.ActiveUIDocument
    app = __revit__.Application
    view = __revit__.ActiveUIDocument.ActiveGraphicalView
    view = doc.ActiveView
    selections = [elId for elId in __revit__.ActiveUIDocument.Selection.GetElementIds() if
                  isinstance(doc.GetElement(elId), Grid)]
    grids = [x for x in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).OfClass(Grid).ToElementIds() if
             x not in selections]
    selections = [doc.GetElement(x) for x in selections]
    grids = [doc.GetElement(x).Name for x in grids]

    if len(selections) == 0:
        MessageBox.Show('Вы должны выбрать хотя бы одну ось', 'Error!')
        raise SystemExit(1)
    res = SelectLevelFrom.show([], title='Перенумерация осей', button_name='Ок')
    if not res:
        raise SystemExit(1)
    if not res['start']:
        raise SystemExit(1)

    PREFIX = res['prefix']
    SUFFIX = res['suffix']
    IS_REVERSE = res['isReverse']

    LENGTH = len(selections)
    NAME = res['start']
    result = []
    IS_DIGIT = NAME.isdigit()

    if not IS_DIGIT:
        CHARACTERS = [
            "А",
            "Б",
            "В",
            "Г",
            "Д",
            "Е",
            "Ж",
            "И",
            "К",
            "Л",
            "М",
            "Н",
            "П",
            "Р",
            "С",
            "Т",
            "У",
            "Ф",
            "Ш",
            "Э",
            "Ю",
            "Я",
        ]
        CHARACTERS_DICT = {
            "А": 1,
            "Б": 2,
            "В": 3,
            "Г": 4,
            "Д": 5,
            "Е": 6,
            "Ж": 7,
            "И": 8,
            "К": 9,
            "Л": 10,
            "М": 11,
            "Н": 12,
            "П": 13,
            "Р": 14,
            "С": 15,
            "Т": 16,
            "У": 17,
            "Ф": 18,
            "Ш": 19,
            "Э": 20,
            "Ю": 21,
            "Я": 22,
        }

        if len(NAME) > 3:
            forms.alert("Некорректно введен номер!", exitscript=True)

        START_NAME = list(NAME)
        for sym in START_NAME:
            if sym not in CHARACTERS_DICT:
                forms.alert("Некорректно введен номер!", exitscript=True)

        START_NUMBER = [CHARACTERS_DICT[x] for x in START_NAME]
        temp = START_NUMBER[::]

        for i in range(LENGTH):
            n = ''
            for index in temp:
                n += CHARACTERS[index - 1]

            result.append(n)
            temp = addChar(temp, CHARACTERS)
    else:
        start = int(NAME)
        for i in range(LENGTH):
            result.append(str(start + i))

    for name in result:
        if PREFIX + name + SUFFIX in grids:
            forms.alert('Имя ' + PREFIX + name + SUFFIX + ' уже занято!', exitscript=True)

    selections = [GridElement(x, IS_DIGIT) for x in selections]
    if len(selections) == 0:
        forms.alert("Вы должны выбрать хотя бы одну ось.", exitscript=True)

    gridElementContainer = GridElementContainer(selections)

    flip = False
    count_flip = 0
    angles_deg = [x.angle for x in selections]
    # Check parrallel
    min_ang = min(angles_deg)
    max_ang = max(angles_deg)
    if max_ang - min_ang > 0.001:
        forms.alert("Линии должны быть параллельны.", exitscript=True)

    # Check flip Condition
    for x in angles_deg:
        if x < 89.999 and x > -0.001 and IS_DIGIT:
            flip = True
            break

    # choose metric to arrang arrays
    angles = [abs(math.cos(x.angle)) for x in selections]
    less_45 = 0
    big_45 = 0
    for ang in angles:
        if ang > 0.707:
            less_45 += 1
        else:
            big_45 += 1

    if big_45 > less_45:
        selections.sort(key=lambda x: x.getRangeX())
    else:
        selections.sort(key=lambda x: x.getRangeY())

    if IS_REVERSE:
        flip = not flip

    if flip:
        result = flipList(result)

    with revit.Transaction("BIM: Нумерация осей"):
        for grid in selections:
            grid.setName(grid.name + 'TEMP')

        for idx, grid in enumerate(selections):
            grid.setName(PREFIX + result[idx] + SUFFIX)

    show_executed_script_notification()


script_execute()
