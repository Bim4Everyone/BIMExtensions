# -*- coding: utf-8 -*-
from pyrevit.framework import List
from pyrevit import revit, DB, UI
from pyrevit import forms
from pyrevit import EXEC_PARAMS

DEBUG = False
MAKELINES = False

selection = revit.get_selection()

import sys
import clr
clr.AddReference('System')
clr.AddReference('System.IO')
clr.AddReference("System.Windows.Forms")

from dosymep_libs.bim4everyone import *

from System.Windows.Forms import MessageBox


def alert(msg):
    MessageBox.Show(msg)


class TransformationMatrix:
    def __init__(self):
        pass


transmatrix = TransformationMatrix()


def createreversedcurve(orig):
    # Create a new curve with the same geometry in the reverse direction.
    if isinstance(orig, DB.Line):
        return DB.Line.CreateBound(orig.GetEndPoint(1), orig.GetEndPoint(0))
    return None


def sortcurvescontiguous(origcurves):
    """
    Sort a list of curves to make them correctly
    ordered and oriented to form a closed loop.
    """
    curves = origcurves
    _inch = 1.0 / 12.0
    _sixteenth = _inch / 16.0
    n = len(curves)
    if DEBUG:
        print('NUMBER OF CURVES: {0}'.format(n))
    # Walk through each curve (after the first)
    # to match up the curves in order
    for i in range(0, n):
        curve = curves[i]
        endpoint = curve.GetEndPoint(1)
        found = (i + 1 >= n)
        for j in range(i + 1, n):
            # If there is a match end->start,
            # this is the next curve
            p = curves[j].GetEndPoint(0)
            if DEBUG:
                print('END2START: {0}'
                      .format(_sixteenth > p.DistanceTo(endpoint)))
            if _sixteenth > p.DistanceTo(endpoint):
                if i + 1 != j:
                    tmp = curves[i + 1]
                    curves[i + 1] = curves[j]
                    curves[j] = tmp
                    if DEBUG:
                        print('SWAPPED.')
                if DEBUG:
                    print('SWAP UNECESSARY.')
                found = True
                break
            # If there is a match end->end,
            # reverse the next curve
            p = curves[j].GetEndPoint(1)
            if DEBUG:
                print('END2END: {0}'
                      .format(_sixteenth > p.DistanceTo(endpoint)))
            if _sixteenth > p.DistanceTo(endpoint):
                if i + 1 == j:
                    curves[i + 1] = createreversedcurve(curves[j])
                else:
                    tmp = curves[i + 1]
                    curves[i + 1] = createreversedcurve(curves[j])
                    curves[j] = tmp
                if DEBUG:
                    print('REVERSED.')
                found = True
                break
        if not found:
            return None
    return curves


def sheet_to_view_transform(sheetcoord):
    global transmatrix
    newx = \
        transmatrix.destmin.X \
        + (((sheetcoord.X - transmatrix.sourcemin.X)
            * (transmatrix.destmax.X - transmatrix.destmin.X))
           / (transmatrix.sourcemax.X - transmatrix.sourcemin.X))

    newy = \
        transmatrix.destmin.Y \
        + (((sheetcoord.Y - transmatrix.sourcemin.Y)
            * (transmatrix.destmax.Y - transmatrix.destmin.Y))
           / (transmatrix.sourcemax.Y - transmatrix.sourcemin.Y))

    return DB.XYZ(newx, newy, 0.0)


@log_plugin(EXEC_PARAMS.command_name)
def script_execute(plugin_logger):
    selview = selvp = None
    vpboundaryoffset = 0.01
    selviewports = []
    selboundary = []

    # pick viewport and line boundary from selection
    for el in selection:
        if isinstance(el, DB.Viewport):
            selviewports.append(el)
        elif isinstance(el, DB.CurveElement):
            selboundary.append(el)
    if len(selviewports) > 0:
        selvp = selviewports[0]
        selview = revit.doc.GetElement(selvp.ViewId)
    else:
        forms.alert('Выберите один вид и замкнутый контур линий детализации. Не выбран вид.', exitscript=True)

    if len(selboundary) < 3:
        forms.alert(
            'Выберите один вид и замкнутый контур линий детализации. Не выбран замкнутый контур линий детализации.', exitscript=True)

    if selview is None:
        forms.alert("Не выбран вид", exitscript=True)

    # making sure the cropbox is active.
    if not selview.CropBoxActive:
        with revit.Transaction('Activate Crop Box'):
            selview.CropBoxActive = True

    # get vp min max points in sheetUCS
    ol = selvp.GetBoxOutline()
    vptempmin = ol.MinimumPoint
    vpmin = DB.XYZ(vptempmin.X + vpboundaryoffset,
                   vptempmin.Y + vpboundaryoffset,
                   0.0)

    vptempmax = ol.MaximumPoint
    vpmax = DB.XYZ(vptempmax.X - vpboundaryoffset,
                   vptempmax.Y - vpboundaryoffset,
                   0.0)
    if DEBUG:
        print('VP MIN MAX: {0}\n'
              '            {1}\n'.format(vpmin, vpmax))

    # get view min max points in modelUCS.
    modelucsx = []
    modelucsy = []
    crsm = selview.GetCropRegionShapeManager()
    cl = crsm.GetCropShape()[0]
    for l in cl:
        modelucsx.append(l.GetEndPoint(0).X)
        modelucsy.append(l.GetEndPoint(0).Y)
        if DEBUG:
            print('CROP LINE POINTS: {0}\n'
                  '                  {1}\n'.format(l.GetEndPoint(0),
                                                   l.GetEndPoint(1)))

    cropmin = DB.XYZ(min(modelucsx), min(modelucsy), 0.0)
    cropmax = DB.XYZ(max(modelucsx), max(modelucsy), 0.0)
    if DEBUG:
        print('CROP MIN MAX: {0}\n'
              '              {1}\n'.format(cropmin, cropmax))
    if DEBUG:
        print('VIEW BOUNDING BOX ON THIS SHEET: {0}\n'
              '                                 {1}\n'
              .format(selview.BoundingBox[selview].Min,
                      selview.BoundingBox[selview].Max))
    transmatrix.sourcemin = vpmin
    transmatrix.sourcemax = vpmax
    transmatrix.destmin = cropmin
    transmatrix.destmax = cropmax

    with revit.Transaction('Set Crop Region'):
        curveloop = []
        for bl in selboundary:
            newlinestart = sheet_to_view_transform(bl.GeometryCurve.GetEndPoint(0))
            newlineend = sheet_to_view_transform(bl.GeometryCurve.GetEndPoint(1))
            geomLine = DB.Line.CreateBound(newlinestart, newlineend)
            if MAKELINES:
                sketchp = selview.SketchPlane
                mline = revit.doc.Create.NewModelCurve(geomLine, sketchp)
            curveloop.append(geomLine)
            if DEBUG:
                print('VP POLY LINE POINTS: {0}\n'
                      '                     {1}\n'
                      .format(bl.GeometryCurve.GetEndPoint(0),
                              bl.GeometryCurve.GetEndPoint(1)))

            if DEBUG:
                print('NEW CROP LINE POINTS: {0}\n'
                      '                      {1}\n'.format(newlinestart,
                                                           newlineend))
        sortedcurves = sortcurvescontiguous(curveloop)
        if sortedcurves:
            crsm.SetCropShape(DB.CurveLoop.Create(List[DB.Curve](sortedcurves)))
        else:
            forms.alert('Контур линий детализации должен быть замкнут.')


if selection:
    script_execute()
else:
    forms.alert('Выберите один вид и замкнутый контур линий детализации.')
