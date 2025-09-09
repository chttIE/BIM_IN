# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import ElementClassFilter, ViewSheet, View, ElementId, ScheduleSheetInstance
from Autodesk.Revit.DB import Transaction, Element
from pyrevit import output, revit, script, forms
output.set_stylesheet(r"Z:\Bim\13. Прочее\pyRevit\outputstyles.css")

output = script.get_output()
lfy = output.linkify
output.close_others(all_open_outputs=True)

doc = __revit__.ActiveUIDocument.Document

from System.Collections.Generic import List

def get_name(el):
    try:
        return Element.Name.GetValue(el)
    except:
        return "-"

def replace(view, doc):
    view_family = str(doc.GetElement(view.GetTypeId()).ViewFamily)

    view_map = {
        "Legend":            ("В_Легенда",          "Легенда"),
        "FloorPlan":         ("В_План этажа",       "План этажа"),
        "CeilingPlan":       ("В_План потолка",     "План потолка"),
        "ThreeDimensional":  ("В_3D-Вид",           "3D-Вид"),
        "Schedule":          ("В_Спецификация",     "Спецификация"),
        "Elevation":         ("В_Фасад",            "Фасад"),
        "Section":           ("В_Разрез",           "Разрез"),
        "Drafting":          ("В_Чертежный вид",    "Чертежный вид"),
    }

    for key in view_map:
        if key in view_family:
            return view_map[key]

    return (":cross_mark: О_Неизвестный", ":cross_mark: Неизвестный вид")

def SetNonControlled(view, view_param):
    view_template = doc.GetElement(view.ViewTemplateId)
    if not view_template:
        return

    current_ids = view_template.GetNonControlledTemplateParameterIds()
    if view_param.Id in current_ids:
        return

    updated_ids = List[ElementId]()
    for eid in current_ids:
        updated_ids.Add(eid)
    updated_ids.Add(view_param.Id)

    view_template.SetNonControlledTemplateParameterIds(updated_ids)
    # output.print_md(">>>:information: Убрал галочку шаблона для параметра *{}*".format(view_param.Definition.Name))

def transfer_parameter_value(doc):
    with Transaction(doc, "Перенос значения параметра") as t:
        t.Start()
        selected_ids = __revit__.ActiveUIDocument.Selection.GetElementIds()
        if not selected_ids:
            print("Выберите хотя бы один элемент!")
            return

        for selected_id in selected_ids:
            view = doc.GetElement(selected_id)
            if not isinstance(view, View):
                print("Элемент {0} не является Видом!".format(view.Name))
                continue

            name_lower = view.Name.lower()
            is_navisworks = "navisworks" in name_lower
            is_cord_plan = "координационный" in name_lower
            # Получаем параметры
            view_param = view.LookupParameter("ADSK_Комплект")
            view_param2 = view.LookupParameter("Назначение вида")

            # Обработка ADSK_Комплект
            if view_param:
                if view_param.IsReadOnly:
                    SetNonControlled(view, view_param)
                if not view_param.IsReadOnly:
                    val = "BIM" if (is_navisworks or is_cord_plan) else "Не размещены на листах"

                    view_param.Set(val)
                else:
                    output.print_md(">>:cross_mark: нет параметра ADSK_Комплект или стоит галочка у шаблона.")
            else:
                output.print_md(">>:cross_mark: Параметр ADSK_Комплект отсутствует у вида {}".format(view.Name))

            # Обработка Назначение вида
            if view_param2:
                if view_param2.IsReadOnly:
                    SetNonControlled(view, view_param2)
                if not view_param2.IsReadOnly:
                    val = "BIM" if (is_navisworks or is_cord_plan) else "В_Рабочие"

                    view_param2.Set(val)
                else:
                    output.print_md(">>:cross_mark: нет параметра Назначение вида или стоит галочка у шаблона.")
            else:
                output.print_md(">>:cross_mark: Параметр Назначение вида отсутствует у вида {}".format(view.Name))

        t.Commit()

# Запуск
transfer_parameter_value(doc)
