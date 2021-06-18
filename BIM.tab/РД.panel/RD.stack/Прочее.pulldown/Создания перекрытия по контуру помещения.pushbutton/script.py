# -*- coding: utf-8 -*-
import os.path as op
import os
import sys
import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import MessageBox
from System.Collections.Generic import List 
from Autodesk.Revit.DB import DisplayUnitType, UnitUtils, IntersectionResultArray, SetComparisonResult, CurveArray, SpatialElementBoundaryOptions, Transform, Arc, Line, ViewDuplicateOption ,ViewType, View, FilteredElementCollector, DimensionType, Transaction, TransactionGroup, ElementId, BuiltInCategory, Grid, InstanceVoidCutUtils, FamilySymbol, FamilyInstanceFilter, Wall, XYZ, WallType
from Autodesk.Revit.DB.Architecture import Room
from Autodesk.Revit.Creation import ItemFactoryBase
import Autodesk
from System.Windows.Forms import MessageBox
from pyrevit import revit
from pyrevit import forms
from pyrevit.framework import Controls
from math import sqrt, acos, asin, sin

__title__ = 'Отделка пола'


doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = __revit__.Application
view = __revit__.ActiveUIDocument.ActiveGraphicalView 
view = doc.ActiveView

class IntersectionRemover:
	def __init__(self, curves):

		self.__curves = curves
		
		self.__contours = self.__splitContours(curves, 0, len(curves)-1)

		self.__contours = self.__cutContours(self.__contours)
		# print "---"
		# for contour in self.__contours:
		# 	print "________"
		# 	for index, curve in enumerate(contour):
		# 		print curve.GetEndPoint(0).IsAlmostEqualTo(contour[index-1].GetEndPoint(1))
		# for cutedContour in cutedContours:
		# 	print len(cutedContour)
		# 	print cutedContour

	def __splitContours(self, curves, startIndex, endIndex):
		contours = []
		contour = []
		index = startIndex
		intersectionResultArray = clr.Reference[Autodesk.Revit.DB.IntersectionResultArray]()
		
		while index <= (endIndex):
			curve = curves[index]
			contour.append(curve)

			secondIndex = index+2
			hasIntersection = False

			while secondIndex <= endIndex:
				if index == startIndex and secondIndex == endIndex:
					secondIndex += 1
					continue
					
				secondCurve = curves[secondIndex]


				intersect = curve.Intersect(secondCurve, intersectionResultArray)
				

				if intersect == SetComparisonResult.Overlap:


					contours = contours + self.__splitContours(curves, index, secondIndex)
					
					# intersectionResultEnumerator = intersectionResultArray.GetEnumerator()
					# for intersectionResult in intersectionResultEnumerator:
					# 	print intersectionResult.XYZPoint 

					hasIntersection = True
					
					break
				
				
				secondIndex += 1
				

			if hasIntersection:
				index = secondIndex
			else:
				index += 1 
		
		contours.append(contour)

		return contours
	
	def __cutContours(self, contours):
		res = []
		for contour in contours:
			res.append(self.__cutContour(contour))
			
		return res

	def __cutContour(self, curves):
		contour = self.__subCutContour(curves)

		endOfContour = contour[-1]
		startOfContour = contour[0]
		pointOfEndContour = endOfContour.GetEndPoint(1)
		pointOfStartContour = startOfContour.GetEndPoint(0)

		if not pointOfEndContour.IsAlmostEqualTo(pointOfStartContour):
			intersectionResultArray = clr.Reference[Autodesk.Revit.DB.IntersectionResultArray]()
			intersect = endOfContour.Intersect(startOfContour, intersectionResultArray)
			contour[-1] = Line.CreateBound(endOfContour.GetEndPoint(0), intersectionResultArray.Item[0].XYZPoint)
			contour[0] = Line.CreateBound(intersectionResultArray.Item[0].XYZPoint, startOfContour.GetEndPoint(1))

		return contour

	def __subCutContour(self, curves):
		res = None
		if len(curves)<2:
			return curves

		leftContour = self.__subCutContour(curves[:len(curves)/2])
		rightContour = self.__subCutContour(curves[len(curves)/2:])

		endOfLeftContour = leftContour[-1]
		startOfRightContour = rightContour[0]
		pointOfLeftContour = endOfLeftContour.GetEndPoint(1)
		pointOfRightContour = startOfRightContour.GetEndPoint(0)
		
		if not pointOfLeftContour.IsAlmostEqualTo(pointOfRightContour):
			intersectionResultArray = clr.Reference[Autodesk.Revit.DB.IntersectionResultArray]()
			intersect = endOfLeftContour.Intersect(startOfRightContour, intersectionResultArray)
			leftContour[-1] = Line.CreateBound(endOfLeftContour.GetEndPoint(0), intersectionResultArray.Item[0].XYZPoint)
			rightContour[0] = Line.CreateBound(intersectionResultArray.Item[0].XYZPoint, startOfRightContour.GetEndPoint(1))
		
		res = leftContour + rightContour
				
		return res

	def __determineLocation(self, point):
		for contour in self.__contours:
			if self.__checkContourPoint(contour, point):
				return contour

		return None

	def __checkContourPoint(self, contour, checkPoint):
		secondCheckPoint = XYZ(checkPoint.X+10000, checkPoint.Y, checkPoint.Z)
		checkingCurve = Line.CreateBound(checkPoint, secondCheckPoint)
		countOfIntersections = 0

		for curve in contour:
			intersectionResult = checkingCurve.Intersect(curve)
			if intersectionResult == SetComparisonResult.Overlap:
				countOfIntersections += 1

		res = countOfIntersections%2 == 1

		return res

	def getContourByPoint(self, point):
		contour = self.__determineLocation(point)

		return contour


class RoomContourman:
	def __init__(self, room):
		self.__room = room

		self.contours = self.__getClosedContours()
		
		self.contours[0] = self.__connectInlineCurves(self.contours[0])
		
		mainContourIntersections = self.__getIntersections(self.contours[0])
		if mainContourIntersections:
			intersectionRemover = IntersectionRemover(self.contours[0])
			self.contours[0] = intersectionRemover.getContourByPoint(room.Location.Point)

	def saveLongestCurve(self):
		curve_with_max_len = None
		max_len = 0
		for curve in self.contours[0]:
			if curve.Length > max_len:
				curve_with_max_len = curve
				max_len = curve.Length

		try:
			self.__room.LookupParameter('Отд_Длинная стена').set(max_len)
		except Exception as e:
			pass

	def getMainContour(self):
		return self.contours[0]

	def getSubContours(self):
		return self.contours[1:]

	def __connectInlineCurves(self, contour):
		res = None
		if len(contour)<2:
			return contour

		leftContour = self.__connectInlineCurves(contour[:len(contour)/2])
		rightContour = self.__connectInlineCurves(contour[len(contour)/2:])

		endOfLeftContour = leftContour[-1]
		startOfRightContour = rightContour[0]
		directionOfLeftContour = endOfLeftContour.Direction
		directionOfRightContour = startOfRightContour.Direction
		
		if directionOfLeftContour.IsAlmostEqualTo(directionOfRightContour):
			
			curve = Line.CreateBound(endOfLeftContour.GetEndPoint(0), startOfRightContour.GetEndPoint(1))

			res = leftContour[:-1] + [curve] + rightContour[1:]
		else:
			
			res = leftContour + rightContour

		return res

	def __getClosedContours(self):
		res = []
		opt = SpatialElementBoundaryOptions()
		segs = self.__room.GetBoundarySegments(opt)
		for boundarySegments in segs:
			curves = self.__getClosedContour(boundarySegments)

			res.append(curves)

		return res

	def __getClosedContour(self, boundarySegments):
		curves = []
		prevCurve = boundarySegments[boundarySegments.Count -1].GetCurve()

		for boundarySegment in boundarySegments:
			curve = boundarySegment.GetCurve()
			
			if not curve.GetEndPoint(0).IsAlmostEqualTo(prevCurve.GetEndPoint(1)):
				curve = Line.CreateBound(prevCurve.GetEndPoint(1), curve.GetEndPoint(1))
			
			prevCurve = curve
			curves.append(curve)
		
		return curves

	def __getIntersections(self, curves):
		intersections = []

		for index in range(0, len(curves)-2):
			curve = curves[index]
			for secondIndex in range(index+2, len(curves)):
				secondCurve = curves[secondIndex]
				if index == 0 and secondIndex == len(curves)-1:
					continue
				intersectionResult = curve.Intersect(secondCurve)
				
				if intersectionResult == SetComparisonResult.Overlap:
					return True

		return False


def getCurveArray(curves):
	curveArray = CurveArray()

	for curve in curves:
		curveArray.Append(curve)

	return curveArray

selection = [ doc.GetElement( elId ) for elId in __revit__.ActiveUIDocument.Selection.GetElementIds() ]

roomSelection = [x for x in selection if isinstance(x, Room)]
roomContourmans = [RoomContourman(x) for x in roomSelection]



tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()

floors = []

for roomContourman in roomContourmans:
	mainContour = roomContourman.getMainContour()
	if mainContour is None:
		roomContourman.floor = None
	else:
		curveArray = getCurveArray(mainContour)
		floor = doc.Create.NewFloor(curveArray, False)
		roomContourman.floor = floor


t.Commit()
tg.Assimilate()

tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()

for roomContourman in roomContourmans:
	if not (roomContourman.floor is None):
		subContours = roomContourman.getSubContours()
		for subContour in subContours:
			curveArray = getCurveArray(subContour)
			roomContourman.floor.Document.Create.NewOpening(roomContourman.floor, curveArray, False)
	

t.Commit()
tg.Assimilate()


from pyrevit import revit
selection = revit.get_selection()
selection.set_to([roomContourman.floor.Id for roomContourman in roomContourmans if not (roomContourman.floor is None)])
