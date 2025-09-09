#coding: utf-8
from pyrevit import output,script
from Autodesk.Revit.DB import FilteredElementCollector as FEC,\
                            BuiltInParameter, BuiltInCategory, WorksharingUtils
from logIN import lg

title = "Оси и уровни"
doc = __revit__.ActiveUIDocument.Document
user = __revit__.Application.Username

output = script.get_output()
output.close_others(all_open_outputs=True)
output.set_title(title)
output.set_width(1500)
lfy = output.linkify
lg(doc, title)

def get_ws_el(el):
    try:return el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString()
    except:return "-"

def lst_len(lst):
    return sum(1 for _ in lst)

def pinned(el):
    if el.Pinned: return "Да"
    else:return "Нет ⛔"

def level_name_check(el,list_name):
    if len(list_name) != 0:
        if el not in list_name:return "Имя не соответствует АР ⛔"
        else:return "+"    
    else:return "Ошибка проверки(нет моделей АР) ⛔"   

def monitoring(el, doc):
    is_monitored = el.IsMonitoringLinkElement()
    if is_monitored:
        monitored_links = el.GetMonitoredLinkElementIds()
        for monitored_element in monitored_links:
            return doc.GetElement(monitored_element).Name.split(":")[0]
    else: return "⛔"


def get_creator(el,doc):
    try:return WorksharingUtils.GetWorksharingTooltipInfo(doc,el.Id).Creator
    except:return "-"

#Кто изменял элементы
def get_lastchanged(el, doc):     
    try:        
        if doc.IsWorkshared:        
            wti = WorksharingUtils.GetWorksharingTooltipInfo(doc, el.Id)
            lastChanger = wti.LastChangedBy 
        else:lastChanger = 'NotWorkshared'
    except:lastChanger = '⛔'        
    return lastChanger


def get_owner(el,doc):
    try: 
        if doc.IsWorkshared:  
            return WorksharingUtils.GetWorksharingTooltipInfo(doc,el.Id).Owner + "⛔" if WorksharingUtils.GetWorksharingTooltipInfo(doc,el.Id).Owner else "-"
        else: return "-"       
    except: return "-"



#Проверка уровней.Начало
def link_level(doc):
    levels = sorted(FEC(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements(), key=lambda level: level.Elevation) 
    cnt_levels = lst_len(levels)
    if cnt_levels != 0:
        levels_data = sorted(
            [
                (
                    el.Name,
                    lfy(el.Id),
                    pinned(el),
                    get_ws_el(el) if get_ws_el(el)[0:11] == "(99)_Уровни" else  get_ws_el(el) + "⛔" ,
                    monitoring(el, doc),
                    get_lastchanged(el, doc),
                    get_creator(el,doc),
                    get_owner(el,doc)
                )
                for el in levels
            ],
            key=lambda x: x[0]
        )
        levels_data = [(i + 1,) + x for i, x in enumerate(levels_data)]

        output.print_table(
            table_data=levels_data,
            title="**Уровни** (кол-во: {})".format(cnt_levels),
            columns=["№", "Имя", "ID", "Закрепление", "Рабочий набор", "Мониторинг","Последний редактор","Автор","Владелец"],
            formats=['', '', '', '', '', '','', '', '']
        )
    else:
        output.print_md("**Уровни в проекте отсутствуют**")
#Проверка уровней.Конец


#Проверка осей.Начало
def link_grids(doc):
    grids = FEC(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
    cnt_grids = lst_len(grids)
    if cnt_grids != 0:
        grids_data = sorted(
            [
                (
                    el.Name,
                    lfy(el.Id),
                    pinned(el),
                    get_ws_el(el) if get_ws_el(el)[0:8] == "(99)_Оси" else  get_ws_el(el) + "⛔",
                    monitoring(el, doc),
                    get_lastchanged(el, doc),
                    get_creator(el,doc),
                    get_owner(el,doc)
                )
                for el in grids
            ],
            key=lambda x: x[0]
        )
        grids_data = [(i + 1,) + x for i, x in enumerate(grids_data)]

        output.print_table(
            table_data=grids_data,
            title="**Оси** (кол-во: {})".format(cnt_grids),
            columns=["№", "Имя", "ID", "Закрепление", "Рабочий набор", "Мониторинг","Последний редактор","Автор","Владелец"],
            formats=['', '', '', '', '', '', '', '', '']
        )
    else:
        output.print_md("**Оси в проекте отсутствуют**")

#main

output.print_md("___")
output.print_md("## ОСИ И УРОВНИ")
output.print_md("___")
link_grids(doc)
link_level(doc)
print('')
