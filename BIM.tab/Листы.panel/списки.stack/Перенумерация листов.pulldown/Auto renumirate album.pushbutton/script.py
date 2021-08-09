# -*- coding: utf-8 -*-

from pySpeech.ViewSheets import *

__title__ = 'Автонумерация'

__doc__ = 'Полностью перенумеровывает альбом, лист которого выделен.\n' \
          'Нумерация начинается с 1.'

order_view = OrderViewSheetModel(DocumentRepository(__revit__))

order_view.LoadViewSheets()
order_view.CheckUniquesNames()
order_view.OrderViewSheets()