# -*- coding: utf-8 -*-
# pylint: skip-file
# by Roman Golev

import Autodesk.Revit.DB as DB
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document  # type: ignore
uidoc = __revit__.ActiveUIDocument         # type: ignore



def get_viewtype(vft_collector):
    # Находим тип семейства вида именно для FloorPlan
    floor_types = [vft for vft in vft_collector.ToElements()
                   if isinstance(vft, DB.ViewFamilyType) and vft.ViewFamily == DB.ViewFamily.FloorPlan]
    if floor_types:
        return floor_types[0]  # первый подходящий тип

    # если не нашли — вернуть None
    return None

def get_categoryID(cat):
    if cat == "BasePoint":
        cat_obj = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_ProjectBasePoint)
    elif cat == "SurveyPoint":
        cat_obj = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_SharedBasePoint)
    elif cat == "SitePoint":
        cat_obj = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_SitePoint)
    elif cat == "Site":
        cat_obj = DB.Category.GetCategory(doc, DB.BuiltInCategory.OST_Site)
    else:
        return None
    return cat_obj.Id if cat_obj else None

def make_active(a_view):
    uidoc.ActiveView = doc.GetElement(a_view.Id)

def find_coord(doc,name):
    coord_views = []
    elems = (DB.FilteredElementCollector(doc)
             .OfClass(DB.View)
             .WhereElementIsNotElementType()
             .ToElements())
    for elem in elems:
        if elem.ViewType == DB.ViewType.FloorPlan and not elem.IsTemplate:
            if name == elem.Name:
                return elem
    return None

def create_coord_plan(doc, vft_id, lvl,name):
    
    
    lvl_id = lvl.Id
    elem = DB.ViewPlan.Create(doc, vft_id, lvl_id)

    #Прописать значение параметров
    p_k = elem.LookupParameter("ADSK_Комплект")
    if p_k: p_k.Set("BIM")
    p_n = elem.LookupParameter("Назначение вида")
    if p_n: p_n.Set("BIM")
    # снять шаблон вида
    par = elem.get_Parameter(DB.BuiltInParameter.VIEW_TEMPLATE)
    if par:
        par.Set(DB.ElementId.InvalidElementId)

    # базовые настройки
    elem.DetailLevel = DB.ViewDetailLevel.Coarse
    elem.DisplayStyle = DB.DisplayStyle.Wireframe
    elem.Discipline = DB.ViewDiscipline.Coordination

    elem.Name = name

    # включить нужные категории (если они есть в проекте)
    for cname in ("Site", "BasePoint", "SurveyPoint"):
        cid = get_categoryID(cname)
        if cid:
            elem.SetCategoryHidden(cid, False)

    # аннотации / обрезка
    elem.AreAnnotationCategoriesHidden = False
    elem.CropBoxActive = False
    elem.CropBoxVisible = False
    ancrop = elem.get_Parameter(DB.BuiltInParameter.VIEWER_ANNOTATION_CROP_ACTIVE)
    if ancrop:
        ancrop.Set(0)

    # диапазон вида: «без ограничений»
    vr = elem.GetViewRange()
    # ВАЖНО: Unlimited нет. Используем InvalidElementId (или PlanViewRange.Unbounded, если доступно)
    invalid = DB.ElementId.InvalidElementId
    vr.SetLevelId(DB.PlanViewPlane.TopClipPlane, invalid)
    vr.SetLevelId(DB.PlanViewPlane.BottomClipPlane, invalid)
    vr.SetLevelId(DB.PlanViewPlane.ViewDepthPlane, invalid)
    elem.SetViewRange(vr)

    return elem

# ---- запуск ----
vft_col = DB.FilteredElementCollector(doc).OfClass(DB.ViewFamilyType)
level = (DB.FilteredElementCollector(doc)
         .OfCategory(DB.BuiltInCategory.OST_Levels)
         .WhereElementIsNotElementType()
         .FirstElement())  # безопаснее, чем [0]

f = level.get_Parameter(DB.BuiltInParameter.LEVEL_ELEV).AsValueString()
name = "ПЭ_({})_Координационный план".format(f)

existing = find_coord(doc,name)
if not existing:
    vf = get_viewtype(vft_col)
    if not vf:
        forms.alert("Не найден ViewFamilyType для Floor Plan.", exitscript=True)

    with DB.Transaction(doc,"Create Coordination Plan") as t:
        t.Start()
        new_plan = create_coord_plan(doc, vf.Id, level,name)

        t.Commit()
        make_active(new_plan)
else:
    make_active(existing)