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
from Autodesk.Revit.DB import Group, ViewSchedule, CopyPasteOptions, ElementTransformUtils, XYZ, UnitUtils, \
    DisplayUnitType, Line, Reference, ReferenceArray, FilteredElementCollector, DimensionType, Options, Transaction, \
    TransactionGroup, ElementId, Wall, ElementIntersectsSolidFilter, ElementIntersectsElementFilter, Solid, \
    SetComparisonResult, CylindricalFace, Grid, CategoryType, LogicalOrFilter, ElementFilter, ElementCategoryFilter, \
    DetailLine
from Autodesk.Revit.Creation import ItemFactoryBase
from Autodesk.Revit.UI.Selection import PickBoxStyle
from Autodesk.Revit.UI import RevitCommandId, PostableCommand, TaskDialog

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from dosymep.Bim4Everyone.Templates import ProjectParameters
from dosymep.Bim4Everyone.SharedParams import SharedParamsConfig
from dosymep.Bim4Everyone.ProjectParams import ProjectParamsConfig

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
GRAD_EPS = math.radians(0.0000001)


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
        angle = abs(a.AngleTo(b))
        # print angle
        return GRAD_EPS > angle or Utils.isEqualAngle(angle, math.pi)

    @staticmethod
    def isNormal(a, b):
        angle = abs(a.AngleTo(b) - math.pi / 2)
        return GRAD_EPS > angle or Utils.isEqualAngle(angle, math.pi)


class CashLine:
    def __init__(self, object):
        self.object = object
        self.line = self.object.Location.Curve
        self.start = self.line.Tessellate()[0]
        self.end = self.line.Tessellate()[1]
        self.id = self.object.Id

        self.direction = self.line.Direction


class CashWall:
    def __init__(self, obj):
        self.object = obj
        self.box = self.object.BoundingBox[view]
        self.start = self.box.Min
        self.end = self.box.Max
        self.id = self.object.Id
        self.pairs = []
        self.pairs.append([self.start, XYZ(self.start.X, self.end.Y, 0)])
        self.pairs.append([self.start, XYZ(self.end.X, self.start.Y, 0)])
        self.pairs.append([self.end, XYZ(self.start.X, self.end.Y, 0)])
        self.pairs.append([self.end, XYZ(self.end.X, self.start.Y, 0)])

        self.curve = self.object.Location.Curve

        self.orientation = obj.Orientation

        geometry = self.object.get_Geometry(geometryOptions)
        self.faces = []
        for geometryElement in geometry:
            if isinstance(geometryElement, Solid):
                for geomFace in geometryElement.Faces:
                    self.faces.append(geomFace)

    def isIntersect(self, line):
        return any([Utils.intersect(line.start, line.end, x[0], x[1]) for x in self.pairs])

    def getNormalReferences(self, line):
        res = []

        for face in self.getNormalFaces(line):
            res.append(face.Reference)

        return res

    def getNormalFaces(self, line):
        res = []

        for face in self.faces:
            if face.Intersect(line.line) == SetComparisonResult.Overlap and Utils.isParallel(face.FaceNormal,
                                                                                             line.direction):
                res.append(face)

        return res

    def getPoint(self):
        geometry = self.object.get_Geometry(geometryOptions)
        self.faces = []
        for geometryElement in geometry:
            if isinstance(geometryElement, Solid):
                return geometryElement.Edges[0].Tessellate()[0]

    def hasCylindricalFace(self):
        for face in self.faces:
            if isinstance(face, CylindricalFace):
                return True
        return False

    def getReferences(self, line):
        return self.getNormalReferences(line)


class CashGrid:
    def __init__(self, obj):
        self.object = obj
        self.id = self.object.Id

        geometry = self.object.Geometry[geometryOptions]
        lineIterator = geometry.GetEnumerator()
        lineIterator.MoveNext()
        self.line = lineIterator.Current

        self.line = obj.Curve
        self.start = self.line.Tessellate()[0]
        self.end = self.line.Tessellate()[1]

        self.orientation = obj.Curve.Direction
        # print self.orientation

        self.parallelWalls = []
        self.distanceValidWalls = []

    def getNormalReferences(self, line):
        res = []
        directionLine = line.direction

        direction = self.line.Direction

        if Utils.isNormal(directionLine, direction):
            res.append(self.line.Reference)

        return res

    def isIntersect(self, line):
        return Utils.intersect(line.start, line.end, self.start, self.end)

    def getReferences(self, line):
        return self.getNormalReferences(line)

    def getNormalFaces(self, line):
        res = []

        return res


def getWallsFromGroup(group):
    res = []
    elementIds = group.GetMemberIds()
    for elementId in elementIds:
        element = doc.GetElement(elementId)
        if isinstance(element, Wall):
            res.append(CashWall(element))
        elif isinstance(element, Group):
            res = res + getWallsFromGroup(element)

    return res


clr.AddReference('ClassLibrary2')
from ClassLibrary2 import UserControl1

a = UserControl1()
result = a.ShowDialog()

if result:
    scale = float(a.scale.split(' ')[1])
# print scale
else:
    raise (SystemExit(1))


# настройка атрибутов
project_parameters = ProjectParameters.Create(__revit__.Application)
project_parameters.SetupRevitParams(doc, ProjectParamsConfig.Instance.CheckIsNormalGrid,
                                         ProjectParamsConfig.Instance.CheckCorrectDistanceGrid)

tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()

selection = [doc.GetElement(i) for i in uidoc.Selection.GetElementIds()]
selectionWalls = [CashWall(i) for i in selection if isinstance(i, Wall)]
selectionGrids = [CashGrid(i) for i in selection if isinstance(i, Grid)]

groups = [i for i in selection if isinstance(i, Group)]
for group in groups:
    selectionWalls = selectionWalls + getWallsFromGroup(group)

checklist_1 = []
checklist_2 = []

for grid in selectionGrids:

    for wall in selectionWalls:
        if not isinstance(wall.curve, Line):
            continue

        if wall not in checklist_1:
            if Utils.isNormal(grid.orientation, wall.orientation):
                checklist_1.append(wall)
                grid.parallelWalls.append(wall)
                try:
                    wall.object.SetParamValue(ProjectParamsConfig.Instance.CheckIsNormalGrid, 1)
                except:
                    pass
            else:
                try:
                    wall.object.SetParamValue(ProjectParamsConfig.Instance.CheckIsNormalGrid, 0)
                except:
                    pass

    for wall in grid.parallelWalls:

        if wall.faces:
            edgePoint = wall.getPoint()
        else:
            edgePoint = wall.curve.Tessellate()[0]

        point = XYZ(edgePoint.X, edgePoint.Y, grid.start.Z)

        distance = abs((grid.end.Y - grid.start.Y) * point.X - (
                    grid.end.X - grid.start.X) * point.Y + grid.end.X * grid.start.Y - grid.end.Y * grid.start.X) / math.sqrt(
            (grid.end.Y - grid.start.Y) ** 2 + (grid.end.X - grid.start.X) ** 2)

        # a = math.sqrt((grid.end.Y - grid.start.Y)**2 + (grid.end.X - grid.start.X)**2)
        # b = math.sqrt((grid.end.Y - point.Y)**2 + (grid.end.X - point.X)**2)
        # c = math.sqrt((grid.start.Y - point.Y)**2 + (grid.start.X - point.X)**2)
        # p = (a+b+c)/2
        # s = math.sqrt(p*(p-a)*(p-b)*(p-c))
        # h = 2*s/a

        distance = UnitUtils.ConvertFromInternalUnits(distance, DisplayUnitType.DUT_MILLIMETERS)

        if wall not in grid.distanceValidWalls:

            if (round(distance, 5)) % scale > 0:
                try:
                    wall.object.SetParamValue(ProjectParamsConfig.Instance.CheckCorrectDistanceGrid, 0)
                except:
                    pass

            else:
                grid.distanceValidWalls.append(wall)
                try:
                    wall.object.SetParamValue(ProjectParamsConfig.Instance.CheckCorrectDistanceGrid, 1)
                except:
                    pass

t.Commit()
tg.Assimilate()
