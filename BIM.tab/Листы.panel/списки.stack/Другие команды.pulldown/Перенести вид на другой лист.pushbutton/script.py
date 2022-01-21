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


from pyrevit import HOST_APP
from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import script


selViewports = []

dest_sheet = forms.select_sheets(title='Select Target Sheets',
                                 button_name='Select Sheets',
                                 multiple=False)

if dest_sheet:
    cursheet = revit.activeview
    if cursheet is None:
        Environment.Message("Нет активного листа")
        Environment.Exit()
    sel = revit.pick_elements()
    if sel is None:
        Environment.Message("Не вибран вид")
        Environment.Exit()
    for el in sel:
        selViewports.append(el)

    if len(selViewports) > 0:
        with revit.Transaction('Move Viewports'):
            for vp in selViewports:
                if isinstance(vp, DB.Viewport):
                    viewId = vp.ViewId
                    vpCenter = vp.GetBoxCenter()
                    vpTypeId = vp.GetTypeId()
                    cursheet.DeleteViewport(vp)
                    nvp = DB.Viewport.Create(revit.doc,
                                             dest_sheet.Id,
                                             viewId,
                                             vpCenter)
                    nvp.ChangeTypeId(vpTypeId)
                elif isinstance(vp, DB.ScheduleSheetInstance):
                    nvp = \
                        DB.ScheduleSheetInstance.Create(
                            revit.doc, dest_sheet.Id, vp.ScheduleId, vp.Point
                            )
                    revit.doc.Delete(vp.Id)
    else:
        forms.alert('At least one viewport must be selected.')
else:
    forms.alert('You must select at least one sheet to add '
                'the selected viewports to.')
