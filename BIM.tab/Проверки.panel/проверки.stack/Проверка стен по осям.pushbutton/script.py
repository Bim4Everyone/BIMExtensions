# -*- coding: utf-8 -*-

import clr

clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep

clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

import math
from abc import abstractmethod

from pyrevit import forms
from pyrevit import script
from pyrevit import revit
from pyrevit import HOST_APP
from pyrevit import EXEC_PARAMS

from dosymep_libs.bim4everyone import *

from System.Collections.Generic import *
from Autodesk.Revit.DB import *

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

geometryOptions = Options()
geometryOptions.View = view
geometryOptions.ComputeReferences = True

EPS = 1E-9
GRAD_EPS = math.radians(0.0000001)


class Utils:
    def __init__(self):
        pass

    @staticmethod
    def det(a, b, c, d):
        return a * d - b * c

    @staticmethod
    def between(a, b, c):
        return min(a, b) <= c + EPS and c <= max(a, b) + EPS

    @staticmethod
    def intersect_1(a, b, c, d):
        if a > b:
            a, b = b, a
        if c > d:
            c, d = d, c

        return max(a, c) <= min(b, d)

    @staticmethod
    def intersect(a, b, c, d):
        a1 = a.Y - b.Y
        b1 = b.X - a.X
        c1 = -a1 * a.X - b1 * a.Y
        a2 = c.Y - d.Y
        b2 = d.X - c.X
        c2 = -a2 * c.X - b2 * c.Y
        zn = Utils.det(a1, b1, a2, b2)
        if zn != 0:
            x = -Utils.det(c1, b1, c2, b2) * 1. / zn
            y = -Utils.det(a1, c1, a2, c2) * 1. / zn
            return Utils.between(a.X, b.X, x) and Utils.between(a.Y, b.Y, y) and Utils.between(c.X, d.X,
                                                                                               x) and Utils.between(c.Y,
                                                                                                                    d.Y,
                                                                                                                    y)
        else:
            return Utils.det(a1, c1, a2, c2) == 0 and Utils.det(b1, c1, b2, c2) == 0 and Utils.intersect_1(a.X, b.X,
                                                                                                           c.X,
                                                                                                           d.X) and Utils.intersect_1(
                a.Y, b.Y, c.Y, d.Y)

    @staticmethod
    def isEqualAngle(a, b):
        return abs(a - b) < GRAD_EPS

    @staticmethod
    def isEqual(a, b):
        return abs(a - b) < EPS

    @staticmethod
    def isParallel(a, b):
        angle = a.AngleTo(b)
        return GRAD_EPS > angle or Utils.isEqualAngle(angle, math.pi)

    @staticmethod
    def isNormal(a, b):
        angle = abs(a.AngleTo(b) - math.pi / 2)
        return GRAD_EPS > angle or Utils.isEqualAngle(angle, math.pi)

    @staticmethod
    def GetDistance(line, point):
        distance = abs((line.End.Y - line.Start.Y) * point.X
                       - (line.End.X - line.Start.X) * point.Y
                       + line.End.X * line.Start.Y
                       - line.End.Y * line.Start.X) \
                   / math.sqrt((line.End.Y - line.Start.Y) ** 2
                               + (line.End.X - line.Start.X) ** 2)

        if HOST_APP.is_newer_than(2021):
            return UnitUtils.ConvertFromInternalUnits(distance, UnitTypeId.Millimeters)
        else:
            return UnitUtils.ConvertFromInternalUnits(distance, DisplayUnitType.DUT_MILLIMETERS)


class CashedReference:
    def __init__(self, reference):
        self.Reference = reference


class CashedElement:
    def __init__(self, element):
        self.Element = element
        self.ElementId = element.Id

        self.BoundingBox = element.BoundingBox[view]
        self.Max = self.BoundingBox.Max
        self.Min = self.BoundingBox.Min

        self.Lines = []
        self.Lines.append([self.Min, XYZ(self.Min.X, self.Max.Y, 0)])
        self.Lines.append([self.Min, XYZ(self.Max.X, self.Min.Y, 0)])
        self.Lines.append([self.Max, XYZ(self.Min.X, self.Max.Y, 0)])
        self.Lines.append([self.Max, XYZ(self.Max.X, self.Min.Y, 0)])

        self.Direction = None

    def IsNormal(self, cashed_line):
        if self.Direction:
            return Utils.isNormal(self.Direction, cashed_line.Direction)

        return False

    @abstractmethod
    def GetNormalFaces(self, cashed_line):
        pass

    @staticmethod
    def GetFacesByGeometry(geometry_instance):
        if isinstance(geometry_instance, Solid):
            return geometry_instance.Faces

        if isinstance(geometry_instance, GeometryElement):
            return [f for g in geometry_instance if isinstance(g, Solid)
                    for f in CashedElement.GetFacesByGeometry(g)]

        return []

    def IsIntersect(self, cashed_line):
        return any([Utils.intersect(cashed_line.Start, cashed_line.End, line[0], line[1]) for line in self.Lines])

    def GetNormalReferences(self, cashed_line):
        return [f.Reference for f in self.GetNormalFaces(cashed_line)]

    @property
    def LocationPoint(self):
        if isinstance(self.Element.Location, LocationPoint):
            return self.Element.Location.Point
        elif isinstance(self.Element.Location, LocationCurve):
            return self.Element.Location.Curve.Tessellate()[0]


class CashedLine(CashedElement):
    def __init__(self, detail_line):
        CashedElement.__init__(self, detail_line)

        self.Curve = None
        self.Direction = None

        self.Start = None
        self.End = None

    def GetNormalFaces(self, cashed_line):
        return []


class CashedGrid(CashedLine):
    def __init__(self, grid):
        CashedLine.__init__(self, grid)

        self.Curve = self.Element.Curve
        self.Direction = self.Curve.Direction

        self.Start = self.Curve.Tessellate()[0]
        self.End = self.Curve.Tessellate()[1]


class CashedDetailLine(CashedLine):
    def __init__(self, detail_line):
        CashedLine.__init__(self, detail_line)

        self.Curve = self.Element.Location.Curve
        self.Direction = self.Curve.Direction

        self.Start = self.Curve.Tessellate()[0]
        self.End = self.Curve.Tessellate()[1]


class CashedWall(CashedElement):
    def __init__(self, wall):
        CashedElement.__init__(self, wall)
        self.Direction = self.Element.Orientation

    def GetFaces(self):
        geometry = self.Element.get_Geometry(geometryOptions)
        return CashedElement.GetFacesByGeometry(geometry)

    def GetNormalFaces(self, cashed_line):
        return [f for f in self.GetFaces()
                if Utils.isParallel(f.FaceNormal, cashed_line.Direction)]


class CashedFamilyInstance(CashedElement):
    def __init__(self, family_instance):
        CashedElement.__init__(self, family_instance)
        self.Direction = self.Element.FacingOrientation

    def GetNormalFaces(self, cashed_line):
        refs = []
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.NotAReference))
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.Left))
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.Right))
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.Front))
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.Back))
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.Bottom))
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.Top))
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.StrongReference))
        refs.extend(self.Element.GetReferences(FamilyInstanceReferenceType.WeakReference))

        planes = [(ref, SketchPlane.Create(doc, ref).GetPlane()) for ref in refs]
        return [CashedReference(ref) for ref, plane in planes
                if Utils.isParallel(plane.Normal, cashed_line.Direction)]

    def IsNormal(self, cashed_line):
        return len(self.GetNormalFaces(cashed_line)) > 0


def get_cashed_elements(selection):
    cashed_elements = []
    for element in selection:
        if isinstance(element, Wall):
            cashed_elements.append(CashedWall(element))
        elif isinstance(element, FamilyInstance):
            cashed_elements.append(CashedFamilyInstance(element))

    groups = [element for element in selection if isinstance(element, Group)]
    for group in groups:
        cashed_elements.extend(get_walls_from_group(group))

    return cashed_elements


def get_walls_from_group(group):
    walls = []

    element_ids = group.GetMemberIds()
    for element_id in element_ids:
        element = doc.GetElement(element_id)
        if isinstance(element, Wall):
            walls.append(CashedWall(element))
        elif isinstance(element, Group):
            walls.extend(get_walls_from_group(element))

    return walls


def get_scale():
    values = ["0.1", "1", "10", "20"]
    response = forms.ask_for_one_item(values, default="0.1", prompt="Точность, мм", title="Задайте точность проверки")

    if response:
        return float(response)

    script.exit()


def check_walls():
    # настройка атрибутов
    project_parameters = ProjectParameters.Create(__revit__.Application)
    project_parameters.SetupRevitParams(doc, ProjectParamsConfig.Instance.CheckIsNormalGrid,
                                        ProjectParamsConfig.Instance.CheckCorrectDistanceGrid)

    selection = uidoc.GetSelectedElements()
    if len(list(selection)) == 0:
        forms.alert("Выберите хотя бы одну ось, стену или колонну.", title="Предупреждение!", footer="dosymep",
                    exitscript=True)

    selection_grids = [CashedGrid(selected_element) for selected_element in selection if
                       isinstance(selected_element, Grid)]
    if len(selection_grids) == 0:
        forms.alert("Выберите хотя бы одну ось.", title="Предупреждение!", footer="dosymep", exitscript=True)

    cashed_elements = get_cashed_elements(selection)
    if len(cashed_elements) == 0:
        forms.alert("Выберите хотя бы одну стену или колонну.", title="Предупреждение!", footer="dosymep",
                    exitscript=True)

    scale = get_scale()
    with revit.Transaction("BIM: Проверка стен"):
        for cashed_element in cashed_elements:
            cashed_element.Element.SetParamValue(ProjectParamsConfig.Instance.CheckIsNormalGrid, "Нет")
            cashed_element.Element.SetParamValue(ProjectParamsConfig.Instance.CheckCorrectDistanceGrid, "Нет")

        for cashed_grid in selection_grids:
            normal_walls = [cashed_element for cashed_element in cashed_elements if
                            cashed_element.IsNormal(cashed_grid)]
            for cashed_element in normal_walls:
                cashed_element.Element.SetParamValue(ProjectParamsConfig.Instance.CheckIsNormalGrid, "Да")

                distance = Utils.GetDistance(cashed_grid, cashed_element.LocationPoint)
                cashed_element.Element.SetParamValue(ProjectParamsConfig.Instance.CheckCorrectDistanceGrid,
                                                     "Да" if round((distance + 0.000000001) % scale, 8) == 0 else "Нет")


@notification()
@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    check_walls()
    show_executed_script_notification()


script_execute()
