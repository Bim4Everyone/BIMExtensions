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


"""Converts selected legend views to detail views."""

from pyrevit.framework import List
from pyrevit import revit, DB, UI
from pyrevit import script
from pyrevit import forms


__title__ = 'Преобразовать легенду в чертежный вид '
__doc__ = 'Преобразовывает выбранные легенды в чертежные виды.'

class CopyUseDestination(DB.IDuplicateTypeNamesHandler):
    def OnDuplicateTypeNamesFound(self, args):
        return DB.DuplicateTypeAction.UseDestinationTypes


def error(msg):
    forms.alert(msg)
    script.exit()


# get a list of selected legends
selection = [x for x in revit.get_selection()
             if x.ViewType == DB.ViewType.Legend]

if len(selection) > 0:
    # get the first style for Drafting views.
    # This will act as the default style
    for type in DB.FilteredElementCollector(revit.doc)\
                  .OfClass(DB.ViewFamilyType):
        if type.ViewFamily == DB.ViewFamily.Drafting:
            draftingViewType = type
            break

    # iterate over interfacetypes legend views
    for srcView in selection:
        print('\nCOPYING {0}'.format(srcView.Name))
        # get legend view elements and exclude non-copyable elements
        viewElements = DB.FilteredElementCollector(revit.doc, srcView.Id)\
                         .ToElements()

        element_list = []
        for el in viewElements:
            if isinstance(el, DB.Element) \
                    and el.Category \
                    and el.Category.Name != 'Legend Components':
                element_list.append(el.Id)
            else:
                print('SKIPPING ELEMENT WITH ID: {0}'.format(el.Id))
        if len(element_list) < 1:
            print('SKIPPING {0}. NO ELEMENTS FOUND.'.format(srcView.Name))
            continue

        # start creating views and copying elements
        with revit.Transaction('Duplicate Legend as Drafting'):
            destView = DB.ViewDrafting.Create(revit.doc, draftingViewType.Id)
            options = DB.CopyPasteOptions()
            options.SetDuplicateTypeNamesHandler(CopyUseDestination())
            copiedElement = \
                DB.ElementTransformUtils.CopyElements(
                    srcView,
                    List[DB.ElementId](element_list),
                    destView,
                    None,
                    options)

            # matching element graphics overrides and view properties
            for dest, src in zip(copiedElement, element_list):
                destView.SetElementOverrides(dest, srcView.GetElementOverrides(src))
            
            
            def lmdname(name): destView.Name = name
        
        
            nw = NameWrap()
            nw.set(srcView.Name, "", "", "_")
            
            num = 1
            while Environment.SafeCall(lmdname, nw.getfullname(), False)[0] == 0:
                num = num + 1
                nw.set(srcView.Name, "", str(num), "_")

            #destView.Name = srcView.Name
            destView.Scale = srcView.Scale
else:
    error('At least one Legend view must be selected.')
