# -*- coding: utf-8 -*-

from pySpeech.ViewSheets import *
from pyrevit import forms

uiDocument = __revit__.ActiveUIDocument
document = uiDocument.Document

result = forms.ask_for_string(
    default='1',
    prompt='Введите число с которого требуется начать нумерацию.',
    title='Автонумерация'
)


if result:
    if result.isdigit():
        if result:
            order_view = OrderViewSheetModel(DocumentRepository(__revit__), int(result))

            order_view.LoadSelectedViewSheets()
            order_view.CheckUniquesNames()

            order_view.OrderViewSheets()
    else:
        print "Было введено не число."