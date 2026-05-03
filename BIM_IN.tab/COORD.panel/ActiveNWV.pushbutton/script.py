#coding: utf-8

__title__ = "Открыть вид\nNavisworks"
__author__ = 'NistratovIlia'
__doc__ = """Делает активным вид Navisworks если найдет такой"""

from pyrevit import script,revit
from Autodesk.Revit.DB import FilteredElementCollector as FEC, View3D

doc = revit.doc
uidoc = __revit__.ActiveUIDocument
output = script.get_output()

def make_active(a):
    uidoc.ActiveView = doc.GetElement(a.Id)
    pass

def get_view(d,name_view):
    for view in FEC(d).OfClass(View3D):
        if name_view == view.Name:
            return view
    return False

view = get_view(doc, "Navisworks")
if view:
    make_active(view)  
