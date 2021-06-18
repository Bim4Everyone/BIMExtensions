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


from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, SpatialElementBoundaryOptions, Transaction, TransactionGroup

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
error_collector = {}

def GroupByParameter(lst, func):
	res = {}
	for el in lst:
		key = func(el)
		if key in res:
			res[key].append(el)
		else:
			res[key] = [el]
	return res

Environment.Message("Данная команда будет доступна после утверждения общих параметров")
Environment.Exit()

class CastRoom:
	ex = 0
	
	def __init__(self, obj):
		self.obj = obj
		self.stage = obj.LookupParameter('Стадия').AsValueString()
		self.id = obj.Id.ToString()
		self.naming = obj.LookupParameter('КГ_Наименование').AsValueString()
		if self.naming == '(нет)':
			error_collector_add(self.id, 'КГ_Наименование')
		
		self.location = obj.LookupParameter('КГ_Корпус.Секция').AsValueString()
		self.group =  obj.LookupParameter('КГ_Имя подгруппы помещений').AsValueString()
		self.type =  obj.LookupParameter('КГ_Тип помещения').AsValueString()
		self.level =  obj.LookupParameter('Уровень').AsValueString()
		if (self.stage == 'Проект' or self.stage == 'Межквартирные перегородки'):
			if self.location == '(нет)':
				#print 'КГ_Корпус.Секция: '+self.stage+'_'+self.level+'_'+self.id
				error_collector_add(self.id, 'КГ_Корпус.Секция')
			if self.group == '(нет)':
				error_collector_add(self.id, 'КГ_Имя подгруппы помещений')
		self.location_short = obj.LookupParameter('КГ_Корпус.Секция короткое').AsString()
		self.group_short = obj.LookupParameter('КГ_Имя подгруппы пом. короткое').AsString()
		self.type_short =  obj.LookupParameter('КГ_Тип помещения короткий').AsString()
		self.level_short = self.level.replace(' этаж', '')
		self.fire_short =  obj.LookupParameter('КГ_Пожарный отсек короткое').AsString()

	def SetSpeech(self):
		self.obj.LookupParameter('Speech_Корпус.Секция короткое').Set(self.location_short)
		self.obj.LookupParameter('Speech_Имя подгруппы пом. короткое').Set(self.group_short)
		self.obj.LookupParameter('Speech_Тип помещения').Set(self.type_short)
		self.obj.LookupParameter('Speech_Этаж').Set(self.level_short)
		if self.fire_short:
			self.obj.LookupParameter('Speech_Пожарный отсек').Set(self.fire_short)
	
	def SetTotalArea(self):
		self.obj.LookupParameter('Speech_Площадь с коэффициентом').Set(self.area_coeff)
		self.obj.LookupParameter('Speech_Площадь округлённая').Set(self.area_rounded)
	
	def Calculate(self, EPS):
		self.area = (round(self.obj.LookupParameter("Площадь").AsDouble()*self.obj.LookupParameter("КГ_Коэф. расчёта площади").AsDouble()*0.3048*0.3048,EPS))/(0.3048*0.3048)
		self.area_kor = (round(self.obj.LookupParameter("Площадь").AsDouble()*self.obj.LookupParameter("КГ_Понижающий коэффициент").AsDouble()*0.3048*0.3048,EPS))/(0.3048*0.3048) 
		self.area_living = (round(self.obj.LookupParameter("Площадь").AsDouble()*self.obj.LookupParameter("КГ_Коэф. расчёта площади").AsDouble()*0.3048*0.3048,EPS))/(0.3048*0.3048)
		self.area_fact = (round(self.obj.LookupParameter("Площадь").AsDouble()*self.obj.LookupParameter("КГ_Коэф. расчёта площади").AsDouble()*0.3048*0.3048,EPS))/(0.3048*0.3048)
		self.area_coeff = round((self.obj.LookupParameter("КГ_Понижающий коэффициент").AsDouble()*self.obj.LookupParameter("Площадь").AsDouble()*0.3048*0.3048),EPS)/(0.3048*0.3048)
		self.area_rounded = round((self.obj.LookupParameter("Площадь").AsDouble()*0.3048*0.3048),EPS)/(0.3048*0.3048)

		
def error_collector_add(id, err):
	if id in error_collector:
		error_collector[id].append(err)
	else:
		error_collector[id] = [err]

def error_finder_ver_2(rooms):
	for location in rooms:
	
		for flat in rooms[location]:
			
			for room in rooms[location][flat]:
				for sec_room in rooms[location][flat]:
					if room.type != sec_room.type:
						room.ex +=1
			if rooms[location][flat][0].ex>0:
				rooms[location][flat].sort(key=lambda k: k.ex)
				check = False
				for idx,room in enumerate(rooms[location][flat]):
					if check:
						#print 'Ну это вообще: '+room.stage+'_'+room.level+'_'+room.id
						error_collector_add(room.id, 'КГ_Тип помещения')
					else:
						if rooms[location][flat][idx+1].ex > room.ex:
							check = True


rooms = [x for x in FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).ToElements()]
opt = SpatialElementBoundaryOptions()

tg = TransactionGroup(doc, "Update")
tg.Start()
t = Transaction(doc, "Update Sheet Parmeters")
t.Start()
del_num = []
for id,room in enumerate(rooms):
	if room.Area>0:
		# Placed if having Area
		#print 'ok!'
		pass
	elif None == room.Location:
		# No Area and No Location => Unplaced
		#print 'NotPlaced. Deleting...'
		doc.Delete(room.Id)
		del_num.insert(0,id)
	else:
		#must be Redundant or NotEnclosed
		segs = room.GetBoundarySegments(opt);
		if (None==segs) or (len(segs)==0):
			#print 'NotEnclosed'
			error_collector_add(room.Id.ToString(), 'Не окружено')
			del_num.insert(0,id)
		else:
			#print 'Redundant'
			error_collector_add(room.Id.ToString(), 'Избыточное')
			del_num.insert(0,id)
t.Commit()
tg.Assimilate()
for num in del_num:
	del rooms[num]
#print len(rooms)

rooms = [ CastRoom(x) for x in rooms]
rooms.sort(key=lambda k: k.location)
rooms = GroupByParameter(rooms, func = lambda x: x.level)
for level in rooms:
	rooms[level] = GroupByParameter(rooms[level], func = lambda x: x.location)
	for location in rooms[level]:
		rooms[level][location] = GroupByParameter(rooms[level][location], func = lambda x: x.group)

for level in rooms:
		error_finder_ver_2(rooms[level])
if len(error_collector)>0:
	print 'Ошибка'
	for id in error_collector:
		print id + ':'
		for error in error_collector[id]:
			print error