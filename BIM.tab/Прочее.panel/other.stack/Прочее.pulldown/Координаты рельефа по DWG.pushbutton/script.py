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

import clr
clr.AddReference('System')
clr.AddReference('System.Windows.Forms')
from System.Windows.Forms import SaveFileDialog, DialogResult 

from Autodesk.Revit.DB import ImportInstance, TextNote, FilteredElementCollector,CurveElement, Arc
from math import sqrt
import codecs

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
view = __revit__.ActiveUIDocument.ActiveGraphicalView 



def get_point(textNote):
	text = textNote.Text
	if text.find(',')>=0:
		parts = text.split(',')
	else:
		parts = text.split('.')
	if len(parts)>2:
		return None
	if any([x.isdigit() for x in parts]):
		box = textNote.BoundingBox[view]
		x = (box.Max.X + box.Min.X)*304.8/2
		y = (box.Max.Y + box.Min.Y)*304.8/2
		z = float('.'.join(parts))*1000
		res = (x,y,z)
		#print '({x},{y},{z})'.format(x = res[0], y = res[1], z = res[2])
		return res
	else:
		return None
		
textNotes = [get_point(x) for x in FilteredElementCollector(doc).OfClass(TextNote).ToElements() if x.OwnerViewId == view.Id]
textNotes = [x for x in textNotes if isinstance(x, tuple)]

class InstancePoint:
	def __init__(self, dot):
		if isinstance(dot, ImportInstance):
			self.x = dot.BoundingBox[view].Max.X*304.8# + dot.BoundingBox[view].Min.X)*304.8/2
			self.y = dot.BoundingBox[view].Max.Y*304.8# + dot.BoundingBox[view].Min.Y)*304.8/2
		else:
			self.x = dot.GeometryCurve.Center.X*304.8
			self.y = dot.GeometryCurve.Center.Y*304.8
		self.z = self.find_elevation()
		
		#print '{x},{y},{z}'.format(x = str(self.x), y = str(self.x), z = str(self.z))
	@property
	def stringPoint(self):
		return '{x},{y},{z}\r\n'.format(x = str(self.x), y = str(self.y), z = str(self.z))
	
	def find_elevation(self):
		if textNotes is None or len(textNotes) < 3:
			Environment.Message("Нет текстовой заметки с координатами или одной из координат")
			Environment.Exit()
		res = textNotes[0][2]
		resRange = self.get_range(textNotes[0])
		for point in textNotes:
			range = self.get_range(point)
			if range<resRange:
				resRange = range
				res = point[2]
		return res
			
			
	def get_range(self, point):
		x = self.x-point[0]
		y = self.y-point[1]
		res = sqrt(x**2 + y**2)
		return res
		
dots = [InstancePoint(x) for x in FilteredElementCollector(doc).OfClass(ImportInstance).ToElements()]
if not dots:
	dots = [InstancePoint(x) for x in FilteredElementCollector(doc).OfClass(CurveElement).ToElements() if isinstance(x.GeometryCurve, Arc)]
fileDialog = SaveFileDialog()
fileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*"
fileDialog.FilterIndex = 1
fileDialog.DefaultExt = "txt"

if fileDialog.ShowDialog() == DialogResult.OK:
	fileName = fileDialog.FileName
	with codecs.open(fileName,'w+',encoding='utf8') as f:
		for point in dots:
			f.write(point.stringPoint)
		
