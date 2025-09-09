# -*- coding: utf-8 -*-
from Autodesk.Revit.DB import ElementClassFilter, ViewSheet, View, ElementId, ScheduleSheetInstance
from Autodesk.Revit.DB import Transaction, Element
from pyrevit import output, revit,script,forms
output.set_stylesheet(r"Z:\Bim\13. Прочее\pyRevit\outputstyles.css")

output = script.get_output()
lfy = output.linkify

output.close_others(all_open_outputs=True)

def get_name(el):
    try:
        return Element.Name.GetValue(el)
    except:
        return "-"

def replace(view, doc):
    """
    Возвращает переопределённое имя типа вида и его описание на основе ViewFamily.
    
    :param view: Элемент вида
    :param doc: Документ Revit
    :return: tuple (internal_name, display_name)
    """
    view_family = str(doc.GetElement(view.GetTypeId()).ViewFamily)

    view_map = {
        "Legend":            ("О_Легенда",          "Легенда"),
        "FloorPlan":         ("О_План этажа",       "План этажа"),
        "CeilingPlan":       ("О_План потолка",     "План потолка"),
        "ThreeDimensional":  ("О_3D-Вид",           "3D-Вид"),
        "Schedule":          ("О_Спецификация",     "Спецификация"),
        "Elevation":         ("О_Фасад",            "Фасад"),
        "Section":           ("О_Разрез",           "Разрез"),
        "Drafting":          ("О_Чертежный вид",    "Чертежный вид"),
        "Detail":            ("О_Узел",             "Узел")
    }

    for key in view_map:
        if key in view_family:
            # Специальная проверка для легенды
            if key == "Legend":
                viewname = view.Name
                if not viewname.startswith("О_"):
                    new_name = "О_{}".format(viewname)
                    try:
                        view.Name = new_name
                    except Exception as e:
                        print("❌ Не удалось переименовать легенду {}: {}".format(viewname, e))
            return view_map[key]

    # По умолчанию, если ничего не найдено
    return (":cross_mark: О_Неизвестный", ":cross_mark: Неизвестный вид")


from System.Collections.Generic import List

def SetNonControlled(view, view_param):
    """
    Добавляет параметр в список неконтролируемых шаблоном параметров.
    Не удаляет уже имеющиеся элементы в списке.

    :param view: Вид, которому назначен шаблон
    :param view_param: Параметр, который нужно отключить от шаблона
    """
    view_template = doc.GetElement(view.ViewTemplateId)
    if not view_template:
        return

    # Получаем текущие неконтролируемые параметры
    current_ids = view_template.GetNonControlledTemplateParameterIds()

    # Если параметр уже есть в списке — ничего не делаем
    if view_param.Id in current_ids:
        return

    # Создаем новый список и добавляем все текущие + новый параметр
    updated_ids = List[ElementId]()
    for eid in current_ids:
        updated_ids.Add(eid)
    updated_ids.Add(view_param.Id)

    # Обновляем шаблон
    view_template.SetNonControlledTemplateParameterIds(updated_ids)
    # output.print_md(">>>:information: Убрал галочку шаблона для параметра *{}*".format(view_param.Definition.Name))


def transfer_parameter_value(doc):
    with Transaction(doc, "Перенос значения параметра") as t:
        t.Start()
        # Получаем все выбранные элементы
        selected_ids = __revit__.ActiveUIDocument.Selection.GetElementIds()
        if not selected_ids:
            print("Выберите хотя бы один элемент!")
            return

        # Обрабатываем каждый выбранный элемент
        for selected_id in selected_ids:
            selected_element = doc.GetElement(selected_id)
            
            if not isinstance(selected_element, ViewSheet):
                print("Элемент {0} не является листом!".format(selected_element.Name))
                continue

            # Получаем параметр ADSK_Комплект с листа
            sheet_param = selected_element.LookupParameter("ADSK_Комплект")
            if not sheet_param:
                print("У листа {0} нет параметра ADSK_Комплект".format(selected_element.Name))
                continue

            # Получаем значение параметра с листа
            sheet_value = sheet_param.AsString()
            if not sheet_value:
                print("Параметр ADSK_Комплект на листе {0} пустой".format(selected_element.Name))
                continue

            lst_views = []
            name_sheet = "{} - {}".format(selected_element.SheetNumber,selected_element.Name)
            sheet_id  = selected_element.GetDependentElements(ElementClassFilter(ScheduleSheetInstance))
            # output.print_md("___")
            # output.print_md("###Распределяю виды на листе {} - ADSK_Комплект - **{}**".format(lfy(selected_element.Id,name_sheet),sheet_value))
         
            for shId in sheet_id:
                she = doc.GetElement(shId)
                sheid = she.ScheduleId
                she = doc.GetElement(sheid)
                # print(lfy(she.Id))
                lst_views.append(she)
            view_id = selected_element.GetAllPlacedViews() #забираем виды
            for v_id in view_id:
                view = doc.GetElement(v_id)
                # print(lfy(view.Id))
                lst_views.append(view)


            # Получаем все виды, размещённые на выбранном листе
            for view in lst_views:
                type_view = get_name(doc.GetElement(view.GetTypeId()))
                new_type_view,type_view = replace(view,doc)
                # output.print_md("> {} - {} - {}".format(lfy(view.Id,view.Name),type_view,new_type_view))

                # Получаем параметр ADSK_Комплект у вида
                view_param = view.LookupParameter("ADSK_Комплект")
                view_param2 = view.LookupParameter("Назначение вида")
                try:
                    if view_param.IsReadOnly:
                        SetNonControlled(view,view_param)
                    if view_param2.IsReadOnly:
                        SetNonControlled(view,view_param2)
                    if view_param and not view_param.IsReadOnly:
                        # Устанавливаем значение из листа в параметр вида
                        view_param.Set(sheet_value)
                        # print("Значение параметра ADSK_Комплект для вида {0} установлено: {1}".format(view.Name, sheet_value))
                    else:
                        output.print_md(">>:cross_mark: нет параметра ADSK_Комплект или стоит галочка у шаблона.")
                    if view_param2 and not view_param2.IsReadOnly:
                        # Устанавливаем значение из листа в параметр вида
                        view_param2.Set(new_type_view)
                        # print("Значение параметра Назначение вида для вида {0} установлено: {1}".format(view.Name, type_view))
                    else:
                        output.print_md(">>:cross_mark:  нет параметра Назначение вида или стоит галочка у шаблона.")
                except Exception as e:
                    print("Не справился с видом {}. Ошибка {}".format(lfy(view.Id),str(e)))
                # print("{0} - {1}".format(new_type_view, sheet_value))
        # Завершаем транзакцию
        t.Commit()

# Вызов функции
doc = __revit__.ActiveUIDocument.Document  # Получение активного документа через pyRevit
transfer_parameter_value(doc)
