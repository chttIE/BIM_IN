# -*- coding: utf-8 -*-
title = 'Открепить элементы' 
from pyrevit import forms,script
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, FilteredElementCollector as FEC, RevitLinkInstance, Transaction
output = script.get_output()
script.get_output().close_others(all_open_outputs=True)
doc = __revit__.ActiveUIDocument.Document  # noqa

notpingrid,notpinlevel,notpinlink = [], [], []

#Получение коллекторов
links = FEC(doc).OfClass(RevitLinkInstance)
levels = FEC(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()
grids = FEC(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
sbasepoint = FEC(doc).OfCategory(BuiltInCategory.OST_SharedBasePoint).WhereElementIsNotElementType().ToElements()[0]
pbasepoint = FEC(doc).OfCategory(BuiltInCategory.OST_ProjectBasePoint).WhereElementIsNotElementType().ToElements()[0]

#Сбор незакрепленных элементов
for grid in grids:
	if grid.Pinned == True:
		notpingrid.append(grid)
for level in levels:
	if level.Pinned == True:
		notpinlevel.append(level)
for link in links:
	if link.Pinned == True:
		notpinlink.append(link)
if pbasepoint.Pinned == True:
		pbasepoint =pbasepoint
else:
    pbasepoint = 2
if sbasepoint.Pinned == True:
		sbasepoint = sbasepoint
else:
    sbasepoint = 2

#Создание формы
select = forms.SelectFromList.show(["Оси",\
                                    "Уровни",\
                                    "Связи",\
                                    "Базовая точка проекта",\
                                    "Точка съемки"] ,
                                    multiselect = True,
                                    title='Выбери тип элементов для открепления',                                                    
                                    width=300,
                                    height=300,
                                    button_name='Выбрать')

#Транзакция
with Transaction(doc,"AutopinElements") as t:
    t.Start()
    if select != None:
        if "Оси" in select:
            output.print_md("Оси:")
            if bool(notpingrid): # Незакрепленные оси
                for el in notpingrid:
                    el.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(0)
                    output.print_md("Ось *{}* (id: *{}*) откреплена!".format(el.Name,output.linkify(el.Id)))
            else:
                output.print_md("Все оси откреплены!")
        if "Уровни" in select:
            output.print_md("Уровни:")
            if bool(notpinlevel): # Незакрепленные уровни
                for el in notpinlevel:
                    el.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(0)
                    output.print_md("Уровень *{}* (id: *{}*) откреплена!".format(el.Name,output.linkify(el.Id)))
            else:
                output.print_md("Все уровни откреплены!")
        if "Связи" in select:
            output.print_md("Связи:")
            if bool(notpinlink): # Незакрепленные связи
                for el in notpinlink:
                    el.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(0)
                    output.print_md("Связь *{}* (id: *{}*) откреплены!".format(el.Name[:el.Name.find(".rvt")],output.linkify(el.Id)))
            else:
                output.print_md("Все связи откреплены!")
        if "Базовая точка проекта" in select:
            if pbasepoint != 2: # Незакрепленные БТП
                pbasepoint.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(0)
                output.print_md("Базовая точка проекта (id: *{}*) откреплена!".format(output.linkify(pbasepoint.Id)))
            else:
                output.print_md("Базовая точка проекта уже откреплена!")
        if "Точка съемки" in select:
            if sbasepoint != 2: # Незакрепленные ТС
                sbasepoint.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(0)
                output.print_md("Точка съемки (id: *{}*) откреплена!".format(output.linkify(sbasepoint.Id)))
            else:
                output.print_md("Точка съемки уже откреплена!")
    else:
        output.print_md("Ничего не выбрано!")
    t.Commit()

