# -*- coding: utf-8 -*-"

from Autodesk.Revit.DB import BuiltInCategory, ElementId, FamilyInstance, Transaction, FilteredElementCollector
from pyrevit import script
from sup import collect_elements_on_view



from Autodesk.Revit.DB import FilteredElementCollector as FEC
 

doc     = __revit__.ActiveUIDocument.Document   # type: ignore
app     = __revit__.Application                 # type: ignore
doc     = __revit__.ActiveUIDocument.Document   # type: ignore
uidoc   = __revit__.ActiveUIDocument            # type: ignore
output = script.get_output()
lfy = output.linkify
output.close_others(True)
name_parameter = "Код_семейства_Revit"

fams = FEC(doc).OfClass(FamilyInstance).WhereElementIsNotElementType().ToElements()

# for el in fams:
# 	print(el.Name)

category_codes = {
    # BuiltInCategory.OST_Abuts: "УС",
    BuiltInCategory.OST_DuctSystem:                 "ВС",
    BuiltInCategory.OST_DuctTerminal: "ВР",
    BuiltInCategory.OST_AnalysisDisplayStyle: "СА",
    BuiltInCategory.OST_AnalysisResults: "РА",
    BuiltInCategory.OST_Areas: "ОБ",
    BuiltInCategory.OST_BridgeBearings: "ПД",
    BuiltInCategory.OST_BridgeCables: "МК",
    BuiltInCategory.OST_BridgeDecks: "МН",
    BuiltInCategory.OST_BridgeFraming: "МР",
    BuiltInCategory.OST_CableTrayFitting: "ФК",
    BuiltInCategory.OST_CableTrayRun: "ПК",
    BuiltInCategory.OST_CableTray: "КЛ",
    BuiltInCategory.OST_Casework: "КР",
    BuiltInCategory.OST_Ceilings: "ПТ",
    BuiltInCategory.OST_Columns: "КН",
    BuiltInCategory.OST_CommunicationDevices: "УС",
    BuiltInCategory.OST_ConduitFitting: "ФТ",
    BuiltInCategory.OST_ConduitRun: "ПП",
    BuiltInCategory.OST_Conduit: "КП",
    BuiltInCategory.OST_Coordination_Model: "МК",
    BuiltInCategory.OST_CurtainGrids: "СШ",
    BuiltInCategory.OST_CurtainWallPanels: "ПШ",
    BuiltInCategory.OST_Curtain_Systems: "СС",
    BuiltInCategory.OST_CurtainWallMullions: "СТ",
    BuiltInCategory.OST_DataDevices: "УП",
    BuiltInCategory.OST_DetailComponents: "ЭД",
    BuiltInCategory.OST_Doors: "ДВ",
    BuiltInCategory.OST_DuctAccessory: "АВ",
    BuiltInCategory.OST_DuctFitting: "ФВ",
    BuiltInCategory.OST_DuctInsulations: "ИВ",
    BuiltInCategory.OST_DuctLinings: "ВВ",
    BuiltInCategory.OST_DuctCurves: "ВП",
    BuiltInCategory.OST_ElectricalCircuit: "ЭЦ",
    BuiltInCategory.OST_ElectricalEquipment: "ЭО",
    BuiltInCategory.OST_ElectricalFixtures: "ЭП",
    BuiltInCategory.OST_Entourage: "ОС",
    BuiltInCategory.OST_FireAlarmDevices: "УПС",
    BuiltInCategory.OST_FlexDuctCurves: "ГВ",
    BuiltInCategory.OST_FlexPipeCurves: "ГТ",
    BuiltInCategory.OST_Floors: "ПЛ",
    BuiltInCategory.OST_Furniture: "МБ",
    BuiltInCategory.OST_FurnitureSystems: "МС",
    BuiltInCategory.OST_GenericModel: "УМ",
    BuiltInCategory.OST_HVAC_Zones: "ЗН",
    BuiltInCategory.OST_LightingDevices: "ОСП",
    BuiltInCategory.OST_LightingFixtures: "ОП",
    BuiltInCategory.OST_Lines: "ЛН",
    BuiltInCategory.OST_Materials: "МТ",
    BuiltInCategory.OST_MechanicalEquipment: "МО",
    BuiltInCategory.OST_Parking: "ПК",
    BuiltInCategory.OST_Parts: "ДТ",
    BuiltInCategory.OST_PipeAccessory: "АТ",
    BuiltInCategory.OST_PipeFitting: "ФТ",
    BuiltInCategory.OST_PipeInsulations: "ИТ",
    BuiltInCategory.OST_PipeCurves: "ТР",
    BuiltInCategory.OST_PipingSystem: "СТП",
    BuiltInCategory.OST_Planting: "ОЗ",
    BuiltInCategory.OST_PlumbingFixtures: "СП",
    BuiltInCategory.OST_ProjectInformation: "ИП",
    BuiltInCategory.OST_Rebar: "СА",
    BuiltInCategory.OST_RvtLinks: "СР",
    BuiltInCategory.OST_Railings: "ПР",
    BuiltInCategory.OST_StairsRailing: "ПР",
    BuiltInCategory.OST_Ramps: "ПН",
    BuiltInCategory.OST_RasterImages: "РИ",
    BuiltInCategory.OST_Roads: "ДР",
    BuiltInCategory.OST_Roofs: "КР",
    BuiltInCategory.OST_Rooms: "КМ",
    BuiltInCategory.OST_SecurityDevices: "УБ",
    BuiltInCategory.OST_ShaftOpening: "ШП",
    BuiltInCategory.OST_Sheets: "ЛТ",
    BuiltInCategory.OST_Site: "ПЛ",
    BuiltInCategory.OST_MEPSpaces: "ПР",
    BuiltInCategory.OST_SpecialityEquipment: "СО",
    BuiltInCategory.OST_Sprinklers: "СК",
    BuiltInCategory.OST_Stairs: "ЛС",
    BuiltInCategory.OST_StairsLandings: "ЛС",
    BuiltInCategory.OST_StairsRuns: "ЛС",
    BuiltInCategory.OST_StairsStringerCarriage: "ЛС",
    BuiltInCategory.OST_StructuralColumns: "КК",
    BuiltInCategory.OST_StructuralFoundation: "СФ",
    BuiltInCategory.OST_StructuralFraming: "СК",
    
    # BuiltInCategory.OST_Structural: "СА",
    BuiltInCategory.OST_StructuralStiffener: "СРЖ",
    BuiltInCategory.OST_StructuralTruss: "СФ",
    BuiltInCategory.OST_TelephoneDevices: "ТУ",
    BuiltInCategory.OST_Topography: "ТП",
    BuiltInCategory.OST_Walls: "СТ",
    BuiltInCategory.OST_Cornices: "СТ",
    BuiltInCategory.OST_Windows: "ОК",
    BuiltInCategory.OST_Wire: "ПР",
}

def get_category_code(element):
    try:
        category = element.Category
        if category is None:
            return None

        cat_id = category.Id.IntegerValue

        # Сопоставление по Id
        for bic, code in category_codes.items():
            if ElementId(bic).IntegerValue == cat_id:
                return code

        return None  # Если не нашли
    except Exception as e:
        print("Ошибка:", e)
        return None




def changes(elements):
    bad= []
    g,i = 0,0
    with Transaction(doc,"Заполнение параметра") as t:

        t.Start()
        try:
            c=0
            for el in elements: 
                p_d = el.LookupParameter(name_parameter)
                text = get_category_code(el)
                if not text:
                    print("ПУСТОЙ УРОВНЬ У ЭЛЕМЕНТА {} ЗАПОЛНИТЕ РУКАМИ! - КАТЕГОРИЯ {}".format(lfy(el.Id),el.Category.Name))
                    continue

                
                if not text: 
                    print("ИМЯ УРОВНЯ НЕДОСТУПНО У ЭЛЕМЕНТА {} ЗАПОЛНИТЕ РУКАМИ!".format(lfy(el.Id)))
                    continue

                    
                if not p_d:
                    print("ПАРАМЕТР ДЛЯ ЗАПОЛНЕНИЯ ОТСУТСТВУЕТ {} ОТРЕДАКТИРУЙТЕ ПАРАМЕТРЫ!".format(lfy(el.Id)))
                    continue
                if p_d.IsReadOnly: 
                    print("ПАРАМЕТР ДЛЯ ЗАПОЛНЕНИЯ ДОСТУПЕН ТОЛЬКО ДЛЯ ЧТЕНИЯ {} ОТРЕДАКТИРУЙТЕ ПАРАМЕТРЫ!".format(lfy(el.Id)))
                    continue
                if p_d.AsString() == text:
                    i+=1
                    continue
                else:
                    p_d.Set(text)
                    g+=1
                    c+=1

            # doc.Regenerate()

        except Exception as e:
            print("ОШИБКУ У {} КАТЕГОРИИ {} ПРОВЕРЬТЕ ЭЛЕМЕНТ И ПАРАМЕТР!".format(lfy(el.Id),el.Category.Name))


        t.Commit()
    return bad,g,i





# Сбор элементов с вида
custom,category = collect_elements_on_view(doc)
output.freeze()
output.print_md("___")

# for el in custom:
#     level = get_level(el)


bad,g,i = changes(custom)
if bad:
    output.print_md("ПРОБЛЕМЫ:")
    for i,b in enumerate(bad):
        print(bad)

output.print_md("Всего элементов {}".format(len(custom)))
if bad:
    output.print_md("У {} Были ошибки ".format(len(bad)))
output.print_md("У {} Было задано ".format(g))
output.print_md("У {} Уже заполнено".format(i))
output.print_md("___")
output.unfreeze()