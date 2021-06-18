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
    def AppData(cls): return cls.os.getenv('APPDATA')

    @classmethod
    def PrgData(cls): return cls.os.getenv('PROGRAMDATA')

    @classmethod
    def PrgFl64(cls): return cls.os.getenv('PROGRAMFILES')

    @classmethod
    def PrgFl32(cls): return cls.os.getenv('PROGRAMFILES(X86)')

    @classmethod
    def UserDir(cls): return cls.os.getenv('USERPROFILE')

    @classmethod
    def UserDoc(cls): return cls.os.path.join(cls.os.getenv('USERPROFILE'), "Documents")    

    @classmethod
    def JoinPath(cls, pathl, pathr): return cls.os.path.join(pathl, pathr)

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

import os.path as op
import os
import sys
import clr
clr.AddReference('System')
clr.AddReference('System.IO')
clr.AddReference("System.Windows.Forms")
from System.IO import FileInfo
from System.Windows.Forms import MessageBox
from System.Collections.Generic import List
from Autodesk.Revit.DB import CurveArray

class ContourExtractor:
	def __init__(self, room):
		self.__object = room

	def getContour(self):
		opt = SpatialElementBoundaryOptions()
		segs = room.GetBoundarySegments(opt)
		boundarySegmentsList = segs
		contours = []
		for boundarySegments in boundarySegmentsList:
			contour.append(self.__getContourFromBoundarySegments(boundarySegments))

	def __getContourFromBoundarySegments(self, boundarySegments):
		contour = []
		for boundarySegment in boundarySegments:
			curve = boundarySegment.GetCurve()
			contour.append(curve)
		return contour


class RoomContour:
	def __init__(self, contours, location):
		self.contours = contours
		self.location = location
		self.floor = None


class Component:
	def toHandle(self, roomContour):
		pass
	
	def toHandle2(self, roomContour):
		pass


class FloorCreaterComponent(Component):
	def toHandle(self, roomContour):
		curveArray = CurveArray()
		for curve in contour:
			curveArray.Append(curve)
		roomContour.floor = doc.Create.NewFloor(curveArray, False)
		
		return True


class BaseDecorator(Component):
	def __init__(self, component):
		self._component = component

	def toHandle(self, roomContour):
		return self._component.toHandle(roomContour)

	def toHandle2(self, roomContour):
		return self._component.toHandle2(roomContour)


class JoinEdgesDecorator(BaseDecorator):
	def toHandle(self, roomContour):
		newContours = []
		for contour in roomContour.contours:
			newContours.append(self.__joinEdges(contour))

		roomContour.contours = newContours
		return self._component.toHandle(roomContour)
	
	def __joinEdges(self, contour):
		newContour = []
		prevCurve = contour[-1]
		for curve in contour:
			if not curve.GetEndPoint(0).IsAlmostEqualTo(prevCurve.GetEndPoint(1)):
				curve = Line.CreateBound(prevCurve.GetEndPoint(1), curve.GetEndPoint(1))
		
			prevCurve = curve
			newContour.append(curve)

		return newContour


class IntersectionRemoverDecorator(BaseDecorator):
	pass


class PickContourDecorator(BaseDecorator):
	pass



import clr
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox
import os.path as op
import datetime
import shutil

pylocalpath = os.path.join(os.getenv('APPDATA'), 'pyrevit\extensions\lib')
sys.path.append(pylocalpath)

from Autodesk.Revit.DB import ModelPathUtils, BasicFileInfo, WorksharingSaveAsOptions
from pySpeech.Forms import InputFormText, ReserveProjectForm

class DocumentReserver:
	def __init__(self, document):
		self._isFormComplete = False
		self._document = document
		self._reserveDirectory = self._getReserveDocumentDirectory()
		if not op.isdir(self._reserveDirectory):
			self._reserveDirectory = None

		self._reservedName = self._getReserveDocumentName()
		self._refreshReservedName()
		if self._isFormComplete:
			centralDocumentPath = self._getCentralDocumentPath()
			reservePath = op.join(self._reserveDirectory, self._reservedName)
			try:
				shutil.copy2(centralDocumentPath, reservePath)
			except:
				info = Environment.GetLastError()
				Environment.Message(info)
				

	def _refreshReservedName(self):
		
		res = ReserveProjectForm.show([], title='Резервное копирование', button_name='Ок', width=350, height=180, folder=self._reserveDirectory)
		# suffix = InputFormText.show([], title='Введите Суффикс', button_name='Ок', width=250, height=130)
		if res:
			self._isFormComplete = True
			suffix = res["suffix"]
			self._reserveDirectory = res["folder"]
			name = op.join(self._reserveDirectory, self._reservedName)
			splitedName = name[:-4]
			
			if not suffix is None:
				name = splitedName + "_{}.rvt".format(suffix)
				splitedName = name[:-4]
			
			start = 1
			while op.isfile(name):
				name = splitedName + "_{}.rvt".format(start)
				start += 1
				
			self._reservedName = op.basename(name)
	
	def _getCentralDocumentPath(self):
		path = self._document.GetWorksharingCentralModelPath()
		centralModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(path)
		centralModelPath = BasicFileInfo.Extract(doc.PathName).CentralPath
		return centralModelPath

	def _getReserveDocumentDirectory(self):
		centralModelPath = self._getCentralDocumentPath()
		centralModelPathDirectory = op.dirname(centralModelPath)
		rootProjectDirectory = op.dirname(centralModelPathDirectory)
		newPathDirectory = op.join(rootProjectDirectory, "4 - Резервные копии")
		return newPathDirectory

	def _getReserveDocumentName(self):
		centralModelPath = self._getCentralDocumentPath()
		baseName = op.basename(centralModelPath)
		today = datetime.date.today()
		datedName = today.strftime("%Y-%m-%d") + "_" + baseName
		return datedName


doc = __revit__.ActiveUIDocument.Document
__title__ = 'Создать\nрезервную копию'

#ReserveProjectForm.show([], title='Резервное копирование', button_name='Ок', width=300, height=160)

if doc.IsWorkshared:
	DocumentReserver(doc)
else:
	MessageBox.Show('Открыт отсоединенный файл!')


