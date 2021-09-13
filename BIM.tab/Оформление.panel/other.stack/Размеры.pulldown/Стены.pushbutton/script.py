# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
from abc import abstractmethod

import clr
import math

clr.AddReference('System')
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
from System.Collections.Generic import List
from Autodesk.Revit.DB import *

from Autodesk.Revit.Creation import ItemFactoryBase
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.UI import RevitCommandId, PostableCommand

from System import Type

'''
print dir(DocumentManager)
doc = DocumentManager.Instance.CurrentDBDocument
uidoc=DocumentManager.Instance.CurrentUIApplication.ActiveUIDocument
uiapp=DocumentManager.Instance.CurrentUIApplication
app = uiapp.Application
view = doc.ActiveView
'''
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
view = __revit__.ActiveUIDocument.ActiveGraphicalView
view = doc.ActiveView

geometryOptions = Options()
geometryOptions.ComputeReferences = True
geometryOptions.View = view

EPS = 1E-9
GRAD_EPS = math.radians(0.01)


class Utils:
    @staticmethod
    def det(a, b, c, d):
        return a * d - b * c

    @staticmethod
    def between(a, b, c):
        return min(a, b) <= c + EPS and c <= max(a, b) + EPS

    @staticmethod
    def intersect_1(a, b, c, d):
        if (a > b):
            a, b = b, a
        if (c > d):
            c, d = d, c
        return max(a, c) <= min(b, d)

    @staticmethod
    def intersect(a, b, c, d):
        A1 = a.Y - b.Y
        B1 = b.X - a.X
        C1 = -A1 * a.X - B1 * a.Y
        A2 = c.Y - d.Y
        B2 = d.X - c.X
        C2 = -A2 * c.X - B2 * c.Y
        zn = Utils.det(A1, B1, A2, B2)
        if (zn != 0):
            x = -Utils.det(C1, B1, C2, B2) * 1. / zn
            y = -Utils.det(A1, C1, A2, C2) * 1. / zn
            return Utils.between(a.X, b.X, x) and Utils.between(a.Y, b.Y, y) and Utils.between(c.X, d.X,
                                                                                               x) and Utils.between(c.Y,
                                                                                                                    d.Y,
                                                                                                                    y)
        else:
            return Utils.det(A1, C1, A2, C2) == 0 and Utils.det(B1, C1, B2, C2) == 0 and Utils.intersect_1(a.X, b.X,
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
    def GetNormalFaces(self, cashedLine):
        pass

    @staticmethod
    def GetFacesByGeometry(geometryInstance):
        if isinstance(geometryInstance, Solid):
            return geometryInstance.Faces

        if isinstance(geometryInstance, GeometryElement):
            return [f for g in geometryInstance if isinstance(g, Solid)
                    for f in g.Faces]

        return []

    def IsIntersect(self, cashedLine):
        return any([Utils.intersect(cashedLine.Start, cashedLine.End, line[0], line[1]) for line in self.Lines])

    def HasCylindricalFace(self):
        return bool([f for f in self.GetFaces() if isinstance(f, CylindricalFace)])

    def GetNormalReferences(self, cashedLine):
        return [f.Reference for f in self.GetNormalFaces(cashedLine)]


class CashedLine(CashedElement):
    def __init__(self, detailLine):
        CashedElement.__init__(self, detailLine)

        self.Curve = self.Element.Location.Curve
        self.Direction = self.Curve.Direction

        self.Start = self.Curve.Tessellate()[0]
        self.End = self.Curve.Tessellate()[1]

    def GetFaces(self):
        return []


class CashedWall(CashedElement):
    def __init__(self, wall):
        CashedElement.__init__(self, wall)

    def GetFaces(self):
        geometry = self.Element.get_Geometry(geometryOptions)
        return CashedElement.GetFacesByGeometry(geometry)

    def GetNormalFaces(self, cashedLine):
        return [f for f in self.GetFaces()
                if Utils.isParallel(f.FaceNormal, cashedLine.Direction)]


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

    def GetNormalFaces(self, cashedLine):
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
                if Utils.isParallel(plane.Normal, cashedLine.Direction)]


def GetCashedElements(compareLine):
    cashed_elements = []
    for element in elements:
        cashedElement = None
        if isinstance(element, Wall):
            cashedElement = CashedWall(element)
        elif isinstance(element, FamilyInstance):
            cashedElement = CashedColumn(element)

        if cashedElement.IsIntersect(compareLine):
            cashed_elements.append(cashedElement)

    return cashed_elements


def GetReferences(compareLine, compareFaces):
    references = ReferenceArray()
    for index, face in enumerate(compareFaces):
        faceDistance = face.Project(compareLine.Start).Distance

        if IsEqualDistance(index, faceDistance, compareLine, compareFaces):
            references.Append(face.Reference)

    return references


def IsEqualDistance(index, faceDistance, compareLine, compareFaces):
    for compareFace in compareFaces[index + 1:]:
        compareFaceDistance = compareFace.Project(compareLine.Start).Distance

        if Utils.isEqual(faceDistance, compareFaceDistance):
            return False

    return True


def GetNormalReferences(compareLine):
    references = []
    for element in elements:
        cashedElement = None
        if isinstance(element, Wall):
            cashedElement = CashedWall(element)
        elif isinstance(element, FamilyInstance):
            cashedElement = CashedColumn(element)

        if cashedElement.IsIntersect(compareLine):
            references.extend(cashedElement.GetNormalReferences(compareLine))

    return references


filterCategories = List[BuiltInCategory]([BuiltInCategory.OST_Walls,
                                          BuiltInCategory.OST_Columns,
                                          BuiltInCategory.OST_StructuralColumns])

elements = FilteredElementCollector(doc, view.Id) \
    .WherePasses(ElementMulticategoryFilter(filterCategories)) \
    .ToElements()

selection = uidoc.Selection.GetElementIds()
detailLines = [doc.GetElement(i) for i in selection if isinstance(doc.GetElement(i), DetailLine)]
elemIds = List[ElementId]()

with Transaction(doc, "Размеры") as transaction:
    transaction.Start()

    for detailLine in detailLines:
        compareLine = CashedLine(detailLine)
        references = GetNormalReferences(compareLine)

        if len(references) > 0:
            array = ReferenceArray()
            for ref in references:
                array.Append(ref)

            line = Line.CreateBound(compareLine.Start, compareLine.End)
            curve = doc.Create.NewDimension(view, line, array)
            doc.Delete(detailLine.Id)

    transaction.Commit()
