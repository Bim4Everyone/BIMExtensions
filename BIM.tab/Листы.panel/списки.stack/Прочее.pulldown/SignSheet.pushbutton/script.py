# -*- coding: utf-8 -*-

import clr
clr.AddReference("dosymep.Revit.dll")
clr.AddReference("dosymep.Bim4Everyone.dll")

import dosymep.Revit
clr.ImportExtensions(dosymep.Revit)
clr.ImportExtensions(dosymep.Bim4Everyone)

from System.Diagnostics import Process


Process.Start(r"T:\Проектный институт\Отдел стандартизации BIM и RD\BIM\3. Семейства\Утвержденная библиотека семейств\022_Подписи")