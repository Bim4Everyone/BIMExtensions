# -*- coding: utf-8 -*-
"""Replaces current selection with elements inside the groups."""

from pyrevit import revit, DB

__title__ = 'Выбрать для стадий'
__doc__ = 'Выделяет все элементы, находящиеся в выбраных группах'

__context__ = 'selection'

doc = __revit__.ActiveUIDocument.Document

selection = revit.get_selection()


filtered_elements = []
for el in selection:
    if isinstance(el, DB.Group):
        for subelid in el.GetMemberIds():
            subel = doc.GetElement(subelid)
            param = subel.LookupParameter("Стадия возведения")
            
            if param is not None and not param.IsReadOnly:
                filtered_elements.append(subelid)

selection.set_to(filtered_elements)
