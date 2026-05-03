# -*- coding: utf-8 -*-
__title__ = 'Создать вид'
__doc__ = """

"""


from Autodesk.Revit.UI import UIDocument
from pyrevit import forms, revit,script
from Autodesk.Revit.DB import XYZ, BuiltInCategory, DisplayStyle, Document, ElementTypeGroup, \
                                Transaction, View3D, ViewDetailLevel,\
                                ViewDuplicateOption, FilteredElementCollector as FEC

uiapp = __revit__ 
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
user = __revit__.Application.Username
output = script.get_output()
script.get_output().close_others(all_open_outputs=True)

docs = app.Documents

def Hidden_connectors(v,fam_doc):
    connector = FEC(fam_doc,v.Id).OfCategory(BuiltInCategory.OST_ConnectorElem).ToElementIds()
    if connector: v.HideElementsTemporary(connector)
    return

def setting_view(v,fam_doc):
    v.AreAnnotationCategoriesHidden = True
    v.AreImportCategoriesHidden = True
    v.ArePointCloudsHidden = True
    v.CropBoxActive = False
    v.CropBoxVisible = False
    v.DetailLevel = ViewDetailLevel.Fine
    v.DisplayStyle = DisplayStyle.Shading
    v.IsSectionBoxActive = False
    try:
        v.ToggleToIsometric()
    

        v.OrientTo(XYZ(2.27101522431622, 
                    2.32807317165811, 
                    -2.34309571423024))
    except:
        print("Не смог включить изометрию")
    v.RevealConstraintsMode = False
    Hidden_connectors(v,fam_doc)

def view_name(doc_f):
    return next((view for view in FEC(doc_f).OfClass(View3D) if view.Name == 'Миниатюра'), False)


def main(fam_doc):
    if fam_doc.IsFamilyDocument:
        cur_uidoc = UIDocument(fam_doc)  
        new_View = view_name(fam_doc)
        with Transaction(fam_doc,'pyIN | Создание вида для сохранения семейства') as t:
            t.Start()
            if new_View == False:
                view_type_3D_id = fam_doc.GetDefaultElementTypeId(ElementTypeGroup.ViewType3D)
                new_view3d     = View3D.CreateIsometric(fam_doc, view_type_3D_id)
                dupop = ViewDuplicateOption.AsDependent
                dupop = ViewDuplicateOption.WithDetailing
                new_View_id = new_view3d.Id
                new_View = fam_doc.GetElement(new_View_id)
                new_View.Name = "Миниатюра"
                setting_view(new_View,fam_doc)
            else: 
                setting_view(new_View,fam_doc)
            t.Commit()
        cur_uidoc.RequestViewChange(new_View)
        output.print_md("> Сделал активным вид '{}'".format(new_View.Name))
        return True
    else:
        output.print_md("- Не семейство ")
        return False


#MAIN

fam_docs = [d for d in revit.docs if d.IsFamilyDocument] 
if not fam_docs:
    script.exit()
for fam_doc in fam_docs:
    path = fam_doc.PathName
    uiapp.OpenAndActivateDocument(path)
    output.print_md(" - Работа с семейством {}".format(fam_doc.Title))
    main(fam_doc)