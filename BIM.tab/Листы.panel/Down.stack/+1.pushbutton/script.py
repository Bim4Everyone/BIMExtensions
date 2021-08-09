# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")

import dosymep
clr.ImportExtensions(dosymep.Revit)

from pySpeech.ViewSheets import renumber

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument
selection = list(__revit__.ActiveUIDocument.GetSelectedElements())

renumber(1, -1, len(selection), "+1")