# -*- coding: utf-8 -*-

import clr
clr.AddReference('System')
clr.AddReference("System.Windows.Forms")

import math
from abc import abstractmethod

from System.Collections.Generic import *
from Autodesk.Revit.DB import *

from pyrevit import revit
from pyrevit import EXEC_PARAMS
from dosymep_libs.bim4everyone import *


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = doc.ActiveView

geometryOptions = Options()
geometryOptions.ComputeReferences = True
geometryOptions.View = view

EPS = 1E-9
GRAD_EPS = math.radians(0.01)

filter_categories = List[BuiltInCategory]([BuiltInCategory.OST_Walls,
                                           BuiltInCategory.OST_Columns,
                                           BuiltInCategory.OST_StructuralColumns])

elements = FilteredElementCollector(doc, view.Id) \
    .WherePasses(ElementMulticategoryFilter(filter_categories)) \
    .ToElements()

selection = uidoc.Selection.GetElementIds()
detail_lines = [doc.GetElement(i) for i in selection if isinstance(doc.GetElement(i), DetailLine)]


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
        return abs((line.End.Y - line.Start.Y) * point.X
                   - (line.End.X - line.Start.X) * point.Y
                   + line.End.X * line.Start.Y
                   - line.End.Y * line.Start.X) \
               / math.sqrt((line.End.Y - line.Start.Y) ** 2
                           + (line.End.X - line.Start.X) ** 2)


class CashedPlane:
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

    @abstractmethod
    def GetFaces(self):
        pass

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


class CashedLine(CashedElement):
    def __init__(self, detail_line):
        CashedElement.__init__(self, detail_line)

        self.Curve = None
        self.Direction = None

        self.Start = None
        self.End = None

    def GetFaces(self):
        pass

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


class CashedColumn(CashedElement):
    def __init__(self, element):
        CashedElement.__init__(self, element)

    def GetFaces(self):
        faces = []

        geometry = self.Element.get_Geometry(geometryOptions)
        for geometryInstance in geometry:
            for symbolGeometryInstance in geometryInstance.GetInstanceGeometry():
                faces.extend(CashedElement.GetFacesByGeometry(symbolGeometryInstance))

        return faces

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
        return [CashedPlane(ref) for ref, plane in planes
                if Utils.isParallel(plane.Normal, cashed_line.Direction)]


def get_normal_references(compare_line):
    references = []
    for element in elements:
        cashed_element = None
        if isinstance(element, Wall):
            cashed_element = CashedWall(element)
        elif isinstance(element, FamilyInstance):
            cashed_element = CashedColumn(element)

        if cashed_element.IsIntersect(compare_line):
            references.extend(cashed_element.GetNormalReferences(compare_line))

    return references


def create_dimensions():
    with revit.Transaction("BIM: Размеры"):
        for selected_line in detail_lines:
            main_line = CashedDetailLine(selected_line)
            normal_references = get_normal_references(main_line)

            if len(normal_references) > 0:
                array = ReferenceArray()
                for normal_ref in normal_references:
                    array.Append(normal_ref)

                new_line = Line.CreateBound(main_line.Start, main_line.End)
                doc.Create.NewDimension(view, new_line, array)
                doc.Delete(selected_line.Id)


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    create_dimensions()


script_execute()