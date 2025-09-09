# -*- coding: utf-8 -*-

from logIN import lg
from Autodesk.Revit.DB import BuiltInParameter, FilteredElementCollector as FEC, RevitLinkInstance
from pyrevit import revit,forms,script,output
title = 'Связи'
output = script.get_output()
script.get_output().close_others(all_open_outputs=True)
output.set_title(title)
output.set_width(1600)
doc = revit.doc  # noqa
lfy = output.linkify

def get_ws(el):
    try: return el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString()
    except: return "-"

def lst_len(lst): 
    return sum(1 for _ in lst)

def pinned(el):
    if el.Pinned: return "Да"
    else: return "⛔ Нет"

def boundaries(el,doc):
    return doc.GetElement(el.GetTypeId()).get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING).AsValueString()
    
def status(el,doc):
    status = str(doc.GetElement(el.GetTypeId()).GetLinkedFileStatus())
    if status == "InClosedWorkset": return "В закрытом рабочем наборе"
    if status == "Unloaded": return "Выгружена"
    if status == "Loaded": return "Загружена"
    else: return status

def site(el,doc):
    status = str(doc.GetElement(el.GetTypeId()).GetLinkedFileStatus())
    site_ = el.Name.split("позиция ")[1]
    if status == "Unloaded" or status == "InClosedWorkset" : return "Связь не загружена"
    if status == "Loaded" and  site_ == "<Не общедоступное>": return "⛔ Площадка не установлена"
    else: return site_

def links(doc):
    
    links = FEC(doc).OfClass(RevitLinkInstance)  
    cnt_links = lst_len(links)
    if cnt_links != 0:
        output.print_md("## ПРОВЕРКА СВЯЗЕЙ (кол-во: {})".format(cnt_links))
        output.print_md("___")
        output.print_md("Модель: **{}**".format(doc.Title))
        output.print_md("___")
        links_data = sorted(
            [
                (
                    el.Name.split(" : ")[0],
                    get_ws(el) if get_ws(el)[0:8] == "(96)_RVT" else "⛔ " + get_ws(el),
                    boundaries(el,doc),
                    pinned(el),
                    status(el,doc),
                    lfy(el.Id),
                    site(el,doc)
                )
                for el in links
            ],
            key=lambda x: x[1], reverse=False
        )
        links_data = [(i + 1,) + x for i, x in enumerate(links_data)]

        output.print_table(
            table_data=links_data,
            columns=["№", "Имя", "Рабочий набор","Границы помещений","Закрепление","Статус", "ID", "Площадка"],
            formats=['', '', '', '', '', '', '', '']
        )
    else:
        forms.alert('Связи в проекте отсутствуют',ok=False, exitscript=True)
lg(doc,title)
links(doc)
print("")