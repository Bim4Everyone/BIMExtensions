# -*- coding: utf-8 -*-
from pyrevit import HOST_APP
from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import script


__doc__ = 'Открывает список листов для выбора листа в который требуется перенести вид. Запустите эту команду и выберите лист назначения'\
          '. Затем выберите интересующий вид и нажмите кнопку "Готово"'
__title__ = 'Перенести вид на другой лист'

selViewports = []

dest_sheet = forms.select_sheets(title='Select Target Sheets',
                                 button_name='Select Sheets',
                                 multiple=False)

if dest_sheet:
    cursheet = revit.activeview
    sel = revit.pick_elements()
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
