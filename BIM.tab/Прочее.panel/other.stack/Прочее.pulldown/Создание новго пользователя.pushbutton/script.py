# -*- coding: utf-8 -*-
#импорты стандартных библиотек 

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

Environment.Message("Данная команда будет доступна после утверждения")
Environment.Exit()


from operator import itemgetter
import sys
#Speech библиотека
#sys.path.insert(0, 'W:\BIM-Ресурсы\pySpeechLib')
from pySpeech.Forms import CopyUserView
from pySpeech.Data import ViewWorker
from pySpeech import alert, AlbumFilter, ClearString, GetNumber, FormNum

__title__ = 'Новый\nПользователь'

class CopyUserWorker(ViewWorker):
	def _setup(self, **kwargs):
		list = self.GetPurposeList()
		res = CopyUserView.show([], title='Введите имя пользователя', button_name='Ок', width=270, height=170)
		
		if res:
			self.GetByUser('*User')
			
			@self.Transaction
			def save():
				try:
					for view in self.user:
						newview = view.Duplicate(self.dupopt)
						nv = self.GetElement(newview)
						self.ResetTemplate(nv)
						name = self.SplitPrefix(view)
						print type(name[1])
						print type(res)
						self.SetAfterCopy(nv, name[1], res)
				except Exception as e:
					print e

obj = CopyUserWorker()