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

"""Converts selected detail views to legend views."""

from pyrevit.framework import List
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms

logger = script.get_logger()


import sys
import clr

clr.AddReference('System')
clr.AddReference('System.IO')
clr.AddReference("System.Windows.Forms")
from System.Windows.Forms import MessageBox

def alert(msg):
    MessageBox.Show(msg)

class NameWrap:

    def __init__(self):
        self.name = ""
        self.pfix = ""
        self.sfix = ""
        self.cell = ""

    def __str__(self): return getfullname()

    def set(self, name):
        if name is None: return
        self.name = name

    def set(self, name, pfix, sfix, cell):        
        if not name is None: self.name = name
        if not pfix is None: self.pfix = pfix
        if not sfix is None: self.sfix = sfix
        if not cell is None: self.cell = cell

    def add(self, name, pfix, sfix):        
        if not name is None: self.name = self.name + name        
        if not pfix is None: self.pfix = self.pfix + pfix        
        if not sfix is None: self.sfix = self.sfix + sfix

    def getname(self): return self.name

    def getpfix(self): return self.pfix

    def getsfix(self): return self.sfix

    def getfullname(self):
        full = self.name                
        if not full is "" and not self.pfix is "": full = self.pfix + self.cell + full
        if not full is "" and not self.sfix is "": full = full + self.cell + self.sfix 
        return full

class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


# get a list of selected drafting views
selection = [el for el in revit.get_selection()
             if isinstance(el, DB.View)
             and el.ViewType == DB.ViewType.DraftingView]


if not len(selection) > 0:
    forms.alert('At least one Drafting view must be selected.')
    script.exit()


# finding first available legend view
baseLegendView = None
for v in DB.FilteredElementCollector(revit.doc).OfClass(DB.View):
    if v.ViewType == DB.ViewType.Legend:
        baseLegendView = v
        break

if not baseLegendView:
    forms.alert('At least one Legend view must exist in the model.')
    script.exit()

# iterate over interfacetypes drafting views
for src_view in selection:
    print('\nCOPYING {0}'.format(src_view.Name))

    # get drafting view elements and exclude non-copyable elements
    view_elements = \
        DB.FilteredElementCollector(revit.doc, src_view.Id).ToElements()

    elements_to_copy = []
    for el in view_elements:
        if isinstance(el, DB.Element) and el.Category:
            elements_to_copy.append(el.Id)
        else:
            logger.debug('Skipping Element with id: {0}'.format(el.Id))
    if len(elements_to_copy) < 1:
        logger.debug('Skipping {0}. No copyable elements where found.'
                     .format(src_view.Name))
        continue

    # start creating views and copying elements
    with revit.Transaction('Duplicate Drafting as Legend'):
        # copying and pasting elements
        dest_view = revit.doc.GetElement(
            baseLegendView.Duplicate(DB.ViewDuplicateOption.Duplicate)
            )

        options = DB.CopyPasteOptions()
        options.SetDuplicateTypeNamesHandler(CopyUseDestination())
        copied_element = \
            DB.ElementTransformUtils.CopyElements(
                src_view,
                List[DB.ElementId](elements_to_copy),
                dest_view,
                None,
                options)

        # matching element graphics overrides and view properties
        for dest, src in zip(copied_element, elements_to_copy):
            dest_view.SetElementOverrides(dest, src_view.GetElementOverrides(src))
        
        def lmdname(name): dest_view.Name = name
        
        
        nw = NameWrap()
        nw.set(src_view.Name, "", "", "_")
        
        num = 1
        while Environment.SafeCall(lmdname, nw.getfullname(), False)[0] == 0:
            num = num + 1
            nw.set(src_view.Name, "", str(num), "_")
        
        
        #dest_view.Name = src_view.Name
        dest_view.Scale = src_view.Scale
