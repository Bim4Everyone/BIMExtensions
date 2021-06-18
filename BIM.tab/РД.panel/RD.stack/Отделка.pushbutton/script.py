# -*- coding: utf-8 -*-

class Environment:

	#std
	import sys
	import traceback    
	import unicodedata
	import os
	#clr    
	import clr
	clr.AddReference('System')
	clr.AddReference('System.IO')    
	clr.AddReference('PresentationCore')
	clr.AddReference("PresentationFramework")
	clr.AddReference("System.Windows")
	clr.AddReference("System.Xaml")
	clr.AddReference("WindowsBase")


	from System.Windows import MessageBox
	#
	@classmethod
	def Exit(cls): cls.sys.exit("Exit current environment!")

	@classmethod
	def GetLastError(cls):
		info = cls.sys.exc_info()
		infos = ''.join(cls.traceback.format_exception(info[0], info[1], info[2]))
		return infos

	@classmethod
	def Message(cls, msg):
		if not msg is None and type(msg) == str:
			cls.MessageBox.Show(msg)
		else:
			cls.MessageBox.Show("Empty argument or non string error message!")

	@classmethod
	def SafeCall(cls, code, tail, show):        
		back = None
		flag = 0
		info = "empty"
		if not code is None:
			try:
				back = code(tail)
				flag = 1
			except:
				exc_t, exc_v, exc_i = sys.exc_info()           
				info = ''.join(cls.traceback.format_exception(exc_t, exc_v, exc_i))         
				if show: cls.Message(info)         
		return [flag, back, info]

Environment.Message("Данная команда временно недоступна!")
Environment.Exit()

import os.path as op
import os
import sys
from math import sqrt

import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
clr.AddReference('System')
from System.Collections.Generic import List 
from Autodesk.Revit.DB import UnitUtils, IntersectionResultArray, SetComparisonResult, CurveArray, SpatialElementBoundaryOptions, Transform, Arc, Line, ViewDuplicateOption ,ViewType, View, FilteredElementCollector, DimensionType, Transaction, TransactionGroup, ElementId, BuiltInCategory, Grid, InstanceVoidCutUtils, FamilySymbol, FamilyInstanceFilter, Wall, XYZ, WallType

from Autodesk.Revit.DB import Level, Phase, FamilyInstance, ModelLine, DisplayUnitType, FilteredElementCollector, BuiltInCategory, BuiltInParameter, Transaction, TransactionGroup,FamilySymbol, SpatialElementBoundaryOptions,ElementId, WallType, ViewSchedule, CopyPasteOptions, ElementTransformUtils, ScheduleFieldType
from pyrevit import revit, DB
from pyrevit import forms
from pyrevit.forms import TemplateUserInputWindow
from pyrevit.framework import Controls



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
			self.__room.LookupParameter('Отд_Длинная стена').Set(max_len)
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



class Option(object):
	def __init__(self, obj, state=False):
		self.state = state
		self.name = obj.Name
		self.elevation = obj.Elevation
		def __nonzero__(self):
			return self.state
		def __str__(self):
			return self.name

class SelectLevelFrom(forms.TemplateUserInputWindow):
	xaml_source = op.join(op.dirname(__file__),'SelectFromCheckboxes.xaml')
	
	def _setup(self, **kwargs):
		self.checked_only = kwargs.get('checked_only', True)
		button_name = kwargs.get('button_name', None)
		if button_name:
			self.select_b.Content = button_name
		
		self.list_lb.SelectionMode = Controls.SelectionMode.Extended
		
		self.Height = 550
		#for i in range(1,4):
		#	self.purpose.AddText(str(i))

		self._verify_context()
		self._list_options()

	def _verify_context(self):
		new_context = []
		for item in self._context:
			if not hasattr(item, 'state'):
				new_context.append(BaseCheckBoxItem(item))
			else:
				new_context.append(item)

		self._context = new_context

	def _list_options(self, checkbox_filter=None):
		if checkbox_filter:
			self.checkall_b.Content = 'Check'
			self.uncheckall_b.Content = 'Uncheck'
			self.toggleall_b.Content = 'Toggle'
			checkbox_filter = checkbox_filter.lower()
			self.list_lb.ItemsSource = \
				[checkbox for checkbox in self._context
				if checkbox_filter in checkbox.name.lower()]
		else:
			self.checkall_b.Content = 'Выделить все'
			self.uncheckall_b.Content = 'Сбросить выделение'
			self.toggleall_b.Content = 'Инвертировать'
			self.list_lb.ItemsSource = self._context

	def _set_states(self, state=True, flip=False, selected=False):
		all_items = self.list_lb.ItemsSource
		if selected:
			current_list = self.list_lb.SelectedItems
		else:
			current_list = self.list_lb.ItemsSource
		for checkbox in current_list:
			if flip:
				checkbox.state = not checkbox.state
			else:
				checkbox.state = state

		# push list view to redraw
		self.list_lb.ItemsSource = None
		self.list_lb.ItemsSource = all_items

	def toggle_all(self, sender, args):
		"""Handle toggle all button to toggle state of all check boxes."""
		self._set_states(flip=True)

	def check_all(self, sender, args):
		"""Handle check all button to mark all check boxes as checked."""
		self._set_states(state=True)

	def uncheck_all(self, sender, args):
		"""Handle uncheck all button to mark all check boxes as un-checked."""
		self._set_states(state=False)

	def check_selected(self, sender, args):
		"""Mark selected checkboxes as checked."""
		self._set_states(state=True, selected=True)

	def uncheck_selected(self, sender, args):
		"""Mark selected checkboxes as unchecked."""
		self._set_states(state=False, selected=True)

	def button_select(self, sender, args):
		"""Handle select button click."""
		if self.checked_only:
			self.response = [x for x in self._context if x.state]
		else:
			self.response = self._context
		self.response = {'level':self.response,
						'eps':int(self.purpose.Text)}
		self.Close()


def fencode(parse):
	hex = [elem.encode("hex") for elem in parse]
	#print hex
	for id,el in enumerate(hex):
		if el[0] == 'a':
			hex[id] = '39' + el[1]
		elif el[0] == 'b':
			hex[id] = '40' + el[1]
		elif el[0] == 'c':
			hex[id] = '41' + el[1]
		elif el[0] == 'd':
			hex[id] = '42' + el[1]
		elif el[0] == 'e':
			hex[id] = '43' + el[1]
		elif el[0] == 'f':
			hex[id] = '44' + el[1]
	return hex

def lookup(obj, name):
	if obj.LookupParameter(name):
		return obj.LookupParameter(name)
	else:
		params = obj.Parameters
		for param in params:
			hex_name = param.Definition.Name
			if hex_name == name:
				return param
	return None

def GetPhaseId(phaseName, doc):
	res = 0
	collector = FilteredElementCollector(doc)
	collector.OfClass(Phase)
	phases = [phase for phase in collector if phase.Name == phaseName]
	if not phases:
		res = -1
	else:
		res = phases[0]
	return res
	
def GroupByParameter(lst, func):
	res = {}
	for el in lst:
		key = func(el)
		#print type(key)
		if key in res:
			res[key].append(el)
		else:
			res[key] = [el]
	return res
	
class FacingRoom(object):
	
	def __init__(self, obj):
		self.__obj = obj
		self.__beton = 0.0
		self.__gaz = 0.0
		self.__karton = 0.0
		self.__kirpich = 0.0
		self.__uteplitel = 0.0
		self.__perimetr = 0.0
		self.__plintus = 0.0
		self.__holes = {}
		self.__holes['Кирпич'] = []
		self.__holes['Бетон'] = []
		self.__holes['Газоблок'] = []
		self.__holes['Гипсокартон'] = []
		self.__holes['Утеплитель'] = []
		
		
		self.height = self.__obj.LookupParameter('Полная высота').AsDouble()
		segs = self.segments
		for segmentList in segs:
			for boundarySegment in segmentList:
				x = (boundarySegment.GetCurve().GetEndPoint(0).X - boundarySegment.GetCurve().GetEndPoint(1).X)
				y = (boundarySegment.GetCurve().GetEndPoint(0).Y - boundarySegment.GetCurve().GetEndPoint(1).Y)
				length = sqrt(x*x+y*y)
				
				wallHoles = []
				if (boundarySegment.ElementId != ElementId.InvalidElementId):
					
					wallId = boundarySegment.ElementId.ToString()
					wall = doc.GetElement(boundarySegment.ElementId)
					#print wallId
					if wallId.ToString() in holes:

						 
						for hole in holes[wallId.ToString()]:
							roomColl = []
							if hole.ToRoom[phase]:
								roomColl.append(hole.ToRoom[phase].Id.ToString())
							if hole.FromRoom[phase]:
								roomColl.append(hole.FromRoom[phase].Id.ToString())
							#print roomColl
							if self.__obj.Id.ToString() in roomColl:
								wallHoles.append(hole)
								
							
					wallTypeId = wall.GetTypeId()
					wallType = doc.GetElement(wallTypeId)
					
					if isinstance(wallType, WallType):
						square = 0.0
						height = wall.LookupParameter('Неприсоединенная высота').AsDouble()
						#print height
						if height <= self.height:
							#square = round(length*height*(0.3048*0.3048),2)/(0.3048*0.3048)
							square = length*height
						else:
							#square = round(length*self.height*(0.3048*0.3048),2)/(0.3048*0.3048)
							square = length*self.height
						
						
						#print square
						
						structure = wallType.GetCompoundStructure()
						if structure is None:
							continue
						self.__perimetr += length 
						#print wallId
						coreIndex = structure.GetFirstCoreLayerIndex()
						coreLayer = structure.GetLayers()[coreIndex]
						coreMaterial = doc.GetElement(coreLayer.MaterialId)
						coreClass = coreMaterial.MaterialClass
						#print coreIndex
						#print coreMaterial.Name
						#print coreMaterial.MaterialClass
						#print wall.Id

						if coreClass == 'Бетон':
							self.__beton += square
							for hole in wallHoles:
								if hole not in self.__holes['Бетон']:
									self.__holes['Бетон'].append(hole)
						elif coreClass == 'Газоблок':
							self.__gaz += square
							for hole in wallHoles:
								if hole not in self.__holes['Газоблок']:
									self.__holes['Газоблок'].append(hole)
						elif coreClass == 'Кирпич':
							self.__kirpich += square
							for hole in wallHoles:
								if hole not in self.__holes['Кирпич']:
									self.__holes['Кирпич'].append(hole)
						elif coreClass == 'Гипсокартон':
							self.__karton += square
							for hole in wallHoles:
								if hole not in self.__holes['Гипсокартон']:
									self.__holes['Гипсокартон'].append(hole)
						elif coreClass == 'Утеплитель':
							self.__uteplitel += square
							for hole in wallHoles:
								if hole not in self.__holes['Утеплитель']:
									self.__holes['Утеплитель'].append(hole)
						#for layer in structure.GetLayers():
						#	material = doc.GetElement(layer.MaterialId)
						#	print material.MaterialClass
							#print dir(layer)
						#	pass
					elif isinstance(wall, ModelLine):
						continue
					else:
						self.__perimetr += length 
						material = ''
						if wallType.GetMaterialIds(False):
							material = doc.GetElement(wallType.GetMaterialIds(False)[0]).MaterialClass
						elif wall.LookupParameter('Материал'):
							material = doc.GetElement(wall.LookupParameter('Материал').AsElementId()).MaterialClass
						#square = round(length*self.height*(0.3048*0.3048),2)/(0.3048*0.3048)
						square = length*self.height
						
						if material == 'Бетон':
							self.__beton += square
						elif material == 'Газоблок':
							self.__gaz += square
						elif material == 'Кирпич':
							self.__kirpich += square
						elif material == 'Гипсокартон':
							self.__karton += square
						elif material == 'Утеплитель':
							self.__uteplitel += square
		self.__plintus = self.__perimetr
		#print self.__holes
		for material in self.__holes:
			for hole in self.__holes[material]:
				#print hole
				square = 0
				holeTypeId = hole.GetTypeId()
				holeType = doc.GetElement(holeTypeId)
				if hole.LookupParameter('Ширина') is None:
					length = holeType.LookupParameter('Ширина').AsDouble()
				else:
					length = hole.LookupParameter('Ширина').AsDouble()
				if hole.LookupParameter('Высота') is None:
					height = holeType.LookupParameter('Высота').AsDouble()
				else:
					height = hole.LookupParameter('Высота').AsDouble()
				#square = round(height*length*(0.3048*0.3048),2)/(0.3048*0.3048)
				if hole.Category.Name == 'Двери':
					self.__plintus -= length
				
				square = height*length
				if material == 'Бетон':
					self.__beton -= square
				elif material == 'Газоблок':
					self.__gaz -= square
				elif material == 'Кирпич':
					self.__kirpich -= square
				elif material == 'Гипсокартон':
					self.__karton -= square
				elif material == 'Утеплитель':
					self.__uteplitel -= square
				
		
		self.__beton = round(self.__beton*(0.3048*0.3048),EPS)/(0.3048*0.3048)
		self.__gaz = round(self.__gaz*(0.3048*0.3048),EPS)/(0.3048*0.3048)
		self.__kirpich = round(self.__kirpich*(0.3048*0.3048),EPS)/(0.3048*0.3048)
		self.__karton = round(self.__karton*(0.3048*0.3048),EPS)/(0.3048*0.3048)
		self.__uteplitel = round(self.__uteplitel*(0.3048*0.3048),EPS)/(0.3048*0.3048)
		#print self.__beton	
	
		self.setOTDParam('Отд_По Бетону', self.__beton)
		#print self.__gaz		
		self.setOTDParam('Отд_По Газоблоку', self.__gaz)
		#print self.__kirpich	
		self.setOTDParam('Отд_По Кирпичу', self.__kirpich)	
		#print self.__karton		
		self.setOTDParam('Отд_По Гипсокартону', self.__karton)
		
		self.setOTDParam('Отд_По Утеплителю', self.__uteplitel)
		
		self.setOTDParam('Отд_Плинтус Длина', self.__plintus)
		
		self.setOTDParam('Отд_Периметр', self.__perimetr)

		RoomContourman(self.__obj).saveLongestCurve()
	
	def setOTDParam(self, paramName, value): 
			
		try:
			self.__obj.LookupParameter(paramName).Set(value)
		except:
			lst = ('Отд_Плинтус Длина', 'Отд_Периметр')
			if paramName in lst:
				scheduleName = '111 - Отделка Периметр'
			else:
				scheduleName = '111 - Отделка По'
			self.getProjectParameters(scheduleName)
			self.__obj.LookupParameter(paramName).Set(value)
	
	
	@property
	def segments(self):
		opt = SpatialElementBoundaryOptions()
		segs = self.__obj.GetBoundarySegments(opt)
		return segs
		
	@staticmethod
	def getProjectParameters(name):
		fileName = 'W:\\BIM-Ресурсы\\Revit - 4 Стандарты\\!ФОП\\Параметры проекта '+app.VersionNumber+'.rvt'
		docFamily = app.OpenDocumentFile(fileName)
		
		schedules  = [x for x in FilteredElementCollector(docFamily).OfClass(ViewSchedule).ToElements() if x.Name == name]

		ids = List[ElementId]([x.Id for x in schedules])
		
		option = CopyPasteOptions()


		ElementTransformUtils.CopyElements(docFamily, ids, doc, None, option)

		schedules  = [x for x in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements() if x.Name == name]
		ids = List[ElementId]([x.Id for x in schedules])
		doc.Delete(ids)


		docFamily.Close(saveModified = False)
		
		schedules  = [x for x in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements() if x.Name.find('КГ (Ключ.) - Тип помещения')>=0]
		el = schedules[0]

		newField = None


		for field in el.Definition.GetSchedulableFields():
			if field.FieldType == ScheduleFieldType.Instance:
				parameterId = field.ParameterId
				fieldName = field.GetName(doc)
				if fieldName == 'КГ_Площадь кв. по ТЗ':
					newField = field
					break

		try:
			el.Definition.AddField(newField)
		except:
			pass


		
	
	

		
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
app = doc.Application

phase = GetPhaseId('Проект', doc)
LEVEL = []
EPS = 0 #число знаков после запятой
lvls = FilteredElementCollector(doc).OfClass(Level)
ops = [Option(x) for x in lvls] 
ops.sort(key=lambda x: x.elevation)
res = SelectLevelFrom.show(ops,
				button_name='Рассчитать')
if res:
	LEVEL = [x.name for x in res['level']]
	EPS = res['eps']
else:
	raise SystemExit(1)




doors = [x for x in FilteredElementCollector(doc).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_Doors).ToElements()]
windows = [x for x in FilteredElementCollector(doc).OfClass(FamilyInstance).OfCategory(BuiltInCategory.OST_Windows).ToElements()]
holes = GroupByParameter(doors+windows, func = lambda x: x.Host.Id.ToString())
doors = GroupByParameter(doors, func = lambda x: x.Host.Id.ToString())
#doors = GroupByParameter(doors, func = lambda x: x.Host.Id.ToString())
#windows = GroupByParameter(windows, func = lambda x: x.Host.Id.ToString())
#print windows


rooms = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()
rooms = list(filter(lambda x: x.LookupParameter('Стадия').AsValueString() == 'Проект', rooms))
rooms = GroupByParameter(rooms, func = lambda x: x.LookupParameter('Уровень').AsValueString())


tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()

for level in LEVEL:
	[FacingRoom(x) for x in rooms[level]]

t.Commit()
tg.Assimilate()

MessageBox.Show('Готово!')
'''
for room in rooms:
	fencode('Материал несущих конструкций')
	segs = room.segments
	#print len(segs[0])
	message = "BoundarySegment";
	for segmentList in segs:
		for boundarySegment in segmentList:
			x = (boundarySegment.GetCurve().GetEndPoint(0).X - boundarySegment.GetCurve().GetEndPoint(1).X)*0.3048
			y = (boundarySegment.GetCurve().GetEndPoint(0).Y - boundarySegment.GetCurve().GetEndPoint(1).Y)*0.3048
			#message += "\nlength: " + str(sqrt(x*x+y*y)
			message += "\nlength: " + str(round(sqrt(x*x+y*y),4)*1000)
			#message += ";\nCurve start point: (" + str(boundarySegment.GetCurve().GetEndPoint(0).X) + "," + str(boundarySegment.GetCurve().GetEndPoint(0).Y) + "," +str(boundarySegment.GetCurve().GetEndPoint(0).Z) + ")"
			#message += ";\nCurve end point: (" + str(boundarySegment.GetCurve().GetEndPoint(1).X) + "," + str(boundarySegment.GetCurve().GetEndPoint(1).Y) + "," + str(boundarySegment.GetCurve().GetEndPoint(1).Z) + ")";
			#message += ";\nDocument path name: " + room.Document.PathName
			if (boundarySegment.ElementId != ElementId.InvalidElementId):
				message += ";\nElement name: " + doc.GetElement(boundarySegment.ElementId).Name
				message += ";\nElement Id: " + doc.GetElement(boundarySegment.ElementId).Id.ToString()
				wall = doc.GetElement(boundarySegment.ElementId)
				wallType = wall.GetTypeId()
				wallType = doc.GetElement(wallType)
				if isinstance(wallType, WallType):
					structure = wallType.GetCompoundStructure()
					#print dir(structure)
					print structure.GetLastCoreLayerIndex()
					print structure.GetFirstCoreLayerIndex()
					for layer in structure.GetLayers():
						material = doc.GetElement(layer.MaterialId)
						print material.MaterialClass
						#print dir(layer)
						pass
				else:
					#wallType.LookupParameter('Материал несущих конструкций').AsValueString()
					for parameter in  wallType.Parameters:
						hex_name = [elem.encode("hex") for elem in parameter.Definition.Name]
						print hex_name
						if hex_name in material_hex:
							print parameter.Definition.Name
							print parameter.AsValueString()
							break
				print "---"
				#print room.Document.GetElement(wall.GetMaterialIds(False)[0]).Name
			message += ";\n"
			
	print message

el = selection[0]
pSet = el.Parameters
for param in pSet:
	name = param.Definition.Name
	print name
	hex_name = [elem.encode("hex") for elem in name]
	print [elem.encode("hex") for elem in 'Отметка Низа']
	print hex_name
	bin_name = ' '.join(format(ord(x), 'b') for x in name)
	print ' '.join(format(ord(x), 'b') for x in 'Отметка Низа')
	print bin_name
	break
'''