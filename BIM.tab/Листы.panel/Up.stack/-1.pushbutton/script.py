# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)

from pySpeech.ViewSheets import renumber

__doc__ = 'Перемещает лист/группу листов вверх в списке альбома.'
__title__ = '-1'

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selection = list(__revit__.ActiveUIDocument.GetSelectedElements())

renumber(1, 1, len(selection), "-1")
