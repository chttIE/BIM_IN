# -*- coding: utf-8 -*-
# Release v1 11/02/23
title = 'Закрепить элементы' 

from pyrevit import forms,script
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, FilteredElementCollector as FEC, RevitLinkInstance, Transaction


output = script.get_output()
script.get_output().close_others(all_open_outputs=True)
doc = __revit__.ActiveUIDocument.Document

def pin(el,status):
    el.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(status)

def lst_len(lst):
    return sum(1 for _ in lst)


#Получение коллекторов
links = FEC(doc).OfClass(RevitLinkInstance)
levels = FEC(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
grids = FEC(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
sbasepoint = FEC(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()[0]
pbasepoint = FEC(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()[0]

#Создание формы
select = forms.SelectFromList.show(["Оси",\
                                    "Уровни",\
                                    "Связи",\
                                    "Базовая точка проекта",\
                                    "Точка съемки"] ,
                                    multiselect = True,
                                    title='Выбери тип элементов',                                                    
                                    width=350,
                                    height=300,
                                    button_name='Выбрать')

#Транзакция
with Transaction(doc,"Автозакрепление") as t:
    t.Start()
    if select != None:
        if "Оси" in select and lst_len(grids) != 0:
            for el in grids:
                pin(el,1)
        if "Уровни" in select and lst_len(levels) != 0:
            for el in levels:            
                pin(el,1)
        if "Связи" in select and lst_len(links) != 0:
            for el in links:                
                pin(el,1)
        if "Базовая точка проекта" in select:
            pin(pbasepoint,1)# Незакрепленные БТП
        if "Точка съемки" in select:
            pin(sbasepoint,1)# Незакрепленные ТС
    else:
        script.exit()
    t.Commit()

