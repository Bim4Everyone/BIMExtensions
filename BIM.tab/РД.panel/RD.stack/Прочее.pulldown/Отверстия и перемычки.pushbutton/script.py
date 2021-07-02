# -*- coding: utf-8 -*-
import clr
clr.AddReference('dosymep.Bim4Everyone.dll')

import dosymep
clr.ImportExtensions(dosymep.Bim4Everyone)

from pyrevit import revit, DB
from Autodesk.Revit.DB import FilteredElementCollector, BuiltInCategory, BuiltInParameter, Transaction, TransactionGroup, FamilySymbol
from dosymep.Bim4Everyone.SharedParams.SharedParamsConfig import Instance

familyName = "АР-Отверстие под инженерные системы"

sizeWidthParamName = Instance.SizeWidth
sizeDepthParamName = Instance.SizeDepth

existsBulkheadParamName = Instance.BulkheadExists
bulkheadLengthParamName = Instance.BulkheadLength
bulkheadDepthParamName = Instance.BulkheadDepth
bulkheadClassParamName = Instance.BulkheadClass

doc = __revit__.ActiveUIDocument.Document
uidoc = __revit__.ActiveUIDocument

familySymbolIds = [x.Family.GetFamilySymbolIds()\
    for x in FilteredElementCollector(doc).OfClass(FamilySymbol).ToElements()\
    if x.Family.Name == familyName] 

transaction = Transaction(doc, "Обновление значений перемычек")
transaction.Start()

try:
    emptyList = []
    for familySymbolId in familySymbolIds:
        for symbolId in familySymbolId:
            familyInstances = DB.FilteredElementCollector(revit.doc)\
                .WherePasses(DB.FamilyInstanceFilter(revit.doc, symbolId))\
                .ToElements()
                           
            for familyInstance in familyInstances:
                exists = familyInstance.GetParamValue(existsBulkheadParamName)
                
                if exists:
                    subComponentId = familyInstance.GetSubComponentIds()
                    subComponent = doc.GetElement(subComponentId[0])
    
                    width = subComponent.GetParamValue(sizeWidthParamName) * 304.8
    
                    subComponentType = doc.GetElement(subComponent.GetTypeId())
    
                    depth = subComponentType.GetParamValue(sizeDepthParamName) * 304.8
                    typeMark = subComponentType.get_Parameter(BuiltInParameter.ALL_MODEL_TYPE_MARK).AsString()
                        
                    familyInstance.SetParamValue(bulkheadLengthParamName, width)
                    familyInstance.SetParamValue(bulkheadDepthParamName, depth)
                    
                    if typeMark is None:
                        emptyList.append(familyInstance)
                    else:
                        familyInstance.SetParamValue(bulkheadClassParamName, typeMark)
    
                else:
                    familyInstance.SetParamValue(bulkheadLengthParamName, "")
                    familyInstance.SetParamValue(bulkheadDepthParamName, "")
                    familyInstance.SetParamValue(bulkheadClassParamName, "")
    
    transaction.Commit()
except Exception as ex:
    print ex
    transaction.RollBack()

if emptyList:
    print "Не заполнен параметр 'Маркировка типоразмера'"
    for index in emptyList:
        print index.Id