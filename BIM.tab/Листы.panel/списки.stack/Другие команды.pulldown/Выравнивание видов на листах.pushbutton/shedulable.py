# -*- coding: utf-8 -*-
import os.path as op
import os
import sys

from pyrevit import revit

from Autodesk.Revit.DB import Wall, GroupType, FilteredElementCollector, Location, Transaction, TransactionGroup, LocationCurve, ViewSchedule, StorageType, Phase,Reference, Options, XYZ, FamilyInstance 
from Autodesk.Revit.UI.Selection import ObjectType, ObjectSnapTypes
doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

SchedulableErrors = []

class BaseSchedulable(object):

	def __init__(self, vs):
		instCollector = FilteredElementCollector(vs.Document, vs.Id)
		instCollector.WhereElementIsNotElementType()
		instCollector = [x for x in instCollector]
		
		self._fields = instCollector
		
		self.table = vs
		
		self.full_name = self.table.Name
		self.name = self.table.Name[self.table.Name.find('КГ (Ключ.)')+13:]

		self._setup()
		
		
		
	def _setup(self):
		pass
		
	@staticmethod
	def string_checker(param):
		if param.HasValue:
			if param.StorageType == StorageType.String:
				value = param.AsString()
			else:
				value = param.AsValueString()
			if len(value)<1:
				return True
			else:
				return False
		else:
			return True
	
	@staticmethod
	def double_checker(param):
		if param.HasValue:
			value = param.AsDouble()
			if value > 0:
				return False
			else:
				return True
		else:
			return True
	
	@staticmethod
	def integer_checker(param):
		if param.HasValue:
			value = param.AsInteger()
			if value > 0 and value < 20:
				return False
			else:
				return True
		else:
			return True
			
	@classmethod
	def numerated_string_checker(cls, param):
		if param.HasValue:
			if param.StorageType == StorageType.String:
				return cls.string_checker(param)
			else:
				return False
		else:
			return True
		
	@staticmethod
	def select_checker(param, lst):
		if param.HasValue:
			value = param.AsString()
			if value in lst:
				return False
			else:
				return True
		else:
			return True
	
	
	def rize_exception(self, param):
		errCollector.add_fields(param.Definition.Name, self.name)
	
		
class RoomNameSchedulable(BaseSchedulable):
	def _setup(self):
		double_name = ['КГ_Понижающий коэффициент']
		skip_name = ['Назначение', 'Рабочий набор', 'Редактирует']
		string_name = ['КГ_Наименование','Ключевое имя','Имя','КГ_Понижающий коэф.']
		for el in self._fields:
			for param in el.GetOrderedParameters():
				name = param.Definition.Name
				if name in double_name:
					if self.double_checker(param):
						self.rize_exception(param)
				elif name in skip_name:
					pass
				elif name == 'КГ_Открытое_Закрытое':
					if self.select_checker(param, ['Открытое', 'Закрытое']):
						self.rize_exception(param)
				elif name == 'КГ_Жилое_Нежилое':
					if self.select_checker(param, ['Нежилое', 'Жилое']):
						self.rize_exception(param)
				elif name in string_name:
					if self.string_checker(param):
						self.rize_exception(param)

class FireSchedulable(BaseSchedulable):

	def _setup(self):
		string_names = ['КГ_Пожарный отсек', 'Ключевое имя', 'КГ_Пожарный отсек короткое']
		skip_name = ['Рабочий набор', 'Редактирует']
		for el in self._fields:
			for param in el.GetOrderedParameters():
				name = param.Definition.Name
				if name in string_names:
					if self.string_checker(param):
						self.rize_exception(param)
				elif name in skip_name:
					pass

class SectionSchedulable(BaseSchedulable):

	def _setup(self):
		integer_names = ['КГ_Номера квартир в Секции']
		string_names = ['Ключевое имя', 'КГ_Корпус.Секция короткое']
		skip_name = ['Рабочий набор', 'Редактирует']
		for el in self._fields:
			for param in el.GetOrderedParameters():
				name = param.Definition.Name
				if name in integer_names:
					if self.integer_checker(param):
						self.rize_exception(param)
				elif name in string_names:
					if self.string_checker(param):
						self.rize_exception(param)
				elif name in skip_name:
					pass
					

class RoomTypeSchedulable(BaseSchedulable):
	
	def _setup(self):
		string_names = ['Ключевое имя','КГ_Тип помещения короткий','КГ_Тип помещения']
		
		skip_name = ['Рабочий набор', 'Редактирует']
		for el in self._fields:
			for param in el.GetOrderedParameters():
				name = param.Definition.Name
				if name in string_names:
					if self.string_checker(param):
						self.rize_exception(param)
				elif name in skip_name:
					pass

class GroupNameSchedulable(BaseSchedulable):

	def _setup(self):
		names = ['КГ_Тип нумерации помещений','КГ_Тип нумерации подгрупп']
		string_names = ['КГ_Имя группы помещений', 'Ключевое имя', 'КГ_Имя подгруппы помещений', 'КГ_Имя подгруппы пом. короткое']
		skip_name = ['Рабочий набор', 'Редактирует']
		for el in self._fields:
			for param in el.GetOrderedParameters():
				name = param.Definition.Name
				if name in names:
					if self.numerated_string_checker(param):
						self.rize_exception(param)
				elif name in string_names:
					if self.string_checker(param):
						self.rize_exception(param)
				elif name in skip_name:
					pass

class ScheduleError(object):
	def __init__(self):
		self.fields = {'Имя подгруппы помещений': [],
				'Наименование пом.': [],
				'Пожарные отсеки': [],
				'Тип Корпуса.Секции': [],
				'Тип помещения': []}
		self.tables = {'Имя подгруппы помещений': True,
				'Наименование пом.': True,
				'Пожарные отсеки': True,
				'Тип Корпуса.Секции': True,
				'Тип помещения': True}

	def check_tables(self, tables):
		res = []
		for table in tables:
			name = table.Name[table.Name.find('КГ (Ключ.)')+13:]
			if name in self.tables:
				self.tables[name] = False
				res.append(table)
		return res
	
	def add_fields(self, field, table_name):
		if field not in self.fields[table_name]:
			self.fields[table_name].append(field)
			
			
names = {'Имя подгруппы помещений': GroupNameSchedulable,
				'Наименование пом.': RoomNameSchedulable,
				'Пожарные отсеки': FireSchedulable,
				'Тип Корпуса.Секции': SectionSchedulable,
				'Тип помещения': RoomTypeSchedulable}

schedules  = [x for x in FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements() if x.Name.find('КГ (Ключ.)')>=0]

errCollector = ScheduleError()
schedules = errCollector.check_tables(schedules)
for el in schedules:
	key = el.Name[el.Name.find('КГ (Ключ.)')+13:]		
	names[key](el)
