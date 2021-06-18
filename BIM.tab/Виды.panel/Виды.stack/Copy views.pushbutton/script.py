# -*- coding: utf-8 -*-

#импорты стандартных библиотек 
from operator import itemgetter
import sys
import os

#Speech библиотека
from pySpeech.Forms import CopyFromView
from pySpeech.Data import ViewWorker
from pySpeech import alert, AlbumFilter, ClearString, GetNumber, FormNum
  
class CopyViewWorker(ViewWorker):
	def _setup(self, **kwargs):
	#Проверяет что выбраны виды
		if self.CheckType() and len(self.selection)>0:
			
			list = self.GetPurposeList()#получаем группы
			res = CopyFromView.show("", title='Введите номер', button_name='Ок', width=450, height=230, list=list)
			
			if not res:
				return
			if res['detail']:
				self.dupopt = self.dupopt.WithDetailing
			else:
				self.dupopt = self.dupopt.Duplicate
			@self.Transaction
			def save():
				for view in self.selection:
					
					newview = view.Duplicate(self.dupopt)
					nv = self.GetElement(newview)
					self.ResetTemplate(nv)
					name = self.SplitPrefix(view)
					self.SetAfterCopy(nv, name[1], res)

obj = CopyViewWorker()