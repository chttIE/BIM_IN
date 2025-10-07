# -*- coding: utf-8 -*-"
from pyrevit import script
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Transaction
from sup import collect_elements_on_view


doc     = __revit__.ActiveUIDocument.Document   # type: ignore
app     = __revit__.Application                 # type: ignore
doc     = __revit__.ActiveUIDocument.Document   # type: ignore
uidoc   = __revit__.ActiveUIDocument            # type: ignore
output = script.get_output()
lfy = output.linkify
output.close_others(True)
name_parameter = "ADSK_Этаж"
name_parameter2 = "ЭТАЖ ПРОВЕРОЧНЫЙ"


def get_level(element):
    level = None
    doc = element.Document

    # Пробуем получить уровень по LevelId
    if hasattr(element, "LevelId"):
        level = doc.GetElement(element.LevelId)
        if level:
            return level

    # Иногда уровень хранится в ReferenceLevel
    if hasattr(element, "ReferenceLevel") and element.ReferenceLevel:
        return element.ReferenceLevel

    # Проверка других свойств, если есть
    if hasattr(element, "Level") and element.Level:
        return element.Level

    param = element.get_Parameter(BuiltInParameter.INSTANCE_SCHEDULE_ONLY_LEVEL_PARAM)
    if param:
        level_id = param.AsElementId()
        if level_id:
            return doc.GetElement(level_id)
    if hasattr(element, "Host") and element.Host:
        # print('HOST {}'.format(element.Host.Name))
        return element.Host



    # Альтернативный способ через параметры
    param = element.get_Parameter(BuiltInParameter.SKETCH_PLANE_PARAM)
    if param:
        level_id = param.AsElementId()
        if level_id:
            return doc.GetElement(level_id)



    param = element.get_Parameter(BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM)
    if param:
        level_id = param.AsElementId()
        if level_id:
            return doc.GetElement(level_id)
    
    param = element.get_Parameter(BuiltInParameter.SCHEDULE_LEVEL_PARAM)
    if param:
        level_id = param.AsElementId()
        if level_id:
            return doc.GetElement(level_id)
    
    param = element.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM)
    if param:
        level_id = param.AsElementId()
        if level_id:
            return doc.GetElement(level_id)
    
    # param = element.get_Parameter(BuiltInParameter.INSTANCE_FREE_HOST_PARAM)
    # if param:
    #     level_id = param.AsElementId()
    #     if level_id:
    #         return doc.GetElement(level_id)

    #Вложденки у перил
    if hasattr(element, "HostRailingId"):

        host = doc.GetElement(element.HostRailingId)
        
        if host:
            if hasattr(host, "LevelId"):
                level = doc.GetElement(host.LevelId)
                if level:
                    return level
                
    #Перила и прочие штуки относящие к лестницам            
    if hasattr(element, "GetStairs"):
        stairs = element.GetStairs()
        if stairs:
            param = stairs.get_Parameter(BuiltInParameter.STAIRS_BASE_LEVEL_PARAM)
            if param:
                level_id = param.AsElementId()
                if level_id:
                    return doc.GetElement(level_id)

    # print("НЕ СМОГ ПОЛУЧИТЬ УРОВЕНЬ {} ЗАПОЛИТЕ РУКАМИ".format(lfy(element.Id)))
    return None

import re
from Autodesk.Revit.DB import StorageType

def get_instance_param_value(el, name_parameter):
    """Возвращает значение параметра экземпляра по имени."""
    p = el.LookupParameter(name_parameter)
    if p and p.StorageType == StorageType.String:
        val = p.AsString()
        if val and val.strip():
            return val.strip()
    return None

def extract_number(s):
    """Извлекает число с плавающей точкой из строки (с запятой или точкой)."""
    match = re.search(r'[-+]?\d+[.,]?\d*', s)
    if match:
        return float(match.group(0).replace(',', '.'))
    return None


def get_adsks_floor_value(level_name,ref_floor_value=10.0):
    """
    Вычисляет значение ADSK_Этаж для элемента по его уровню.
    Возвращает число, строку (Roof), или None, если исключён.
    """
    # 1. Проверка: есть ли уже значение в ADSK_Этаж


    # 2. Получаем имя уровня

    if not level_name:
        return None

    # 3. Разбиваем по "_"
    parts = level_name.split("_")
    if len(parts) < 3:
        return level_name  # необычный формат — просто вернём как есть

    # 4. Вытаскиваем отметку (между 2 и 3 подчёркиванием)
    floor_mark = parts[2]  # например: "02.1" или "10" или "Roof"
    try:
        floor_number = extract_number(floor_mark)
    except:
        floor_number = None

    if floor_number is not None:
        # сравниваем с 10-м этажом
        if floor_number < ref_floor_value:
            return floor_number  # например, -1, 2.1 и т.п.
        else:
            return int(floor_number) if floor_number.is_integer() else floor_number
    else:
        # нет числа — значит текст типа Roof, тех. эт
        return parts[3] if len(parts) > 3 else floor_mark


def changes(elements):
    bad= []
    g,i = 0,0
    with Transaction(doc,"Заполнение параметра") as t:

        t.Start()
        try:
            c=0
            for el in elements: 
                current_value = get_instance_param_value(el, "ADSK_Этаж")
                if current_value:
                    continue  # уже заполнено — исключаем
                
                p_d = el.LookupParameter(name_parameter)
 
                level = get_level(el)
                if not level:
                    bad.append("ПУСТОЙ УРОВНЬ У ЭЛЕМЕНТА {} ЗАПОЛНИТЕ РУКАМИ!".format(lfy(el.Id)))
                    continue
                text = level.Name
                
                if not text: 
                    bad.append("ИМЯ УРОВНЯ НЕДОСТУПНО У ЭЛЕМЕНТА {} ЗАПОЛНИТЕ РУКАМИ!".format(lfy(el.Id)))
                    continue

                
                text = get_adsks_floor_value(text)
                
                if not p_d:
                    bad.append("ПАРАМЕТР ДЛЯ ЗАПОЛНЕНИЯ ОТСУТСТВУЕТ {} ОТРЕДАКТИРУЙТЕ ПАРАМЕТРЫ!".format(lfy(el.Id)))
                    continue
                if p_d.IsReadOnly: 
                    bad.append("ПАРАМЕТР ДЛЯ ЗАПОЛНЕНИЯ ДОСТУПЕН ТОЛЬКО ДЛЯ ЧТЕНИЯ {} ОТРЕДАКТИРУЙТЕ ПАРАМЕТРЫ!".format(lfy(el.Id)))
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
            bad.append("ОШИБКУ У {} КАТЕГОРИИ {} ПРОВЕРЬТЕ ЭЛЕМЕНТ И ПАРАМЕТР!".format(lfy(el.Id),el.Category.Name))


        t.Commit()
    return bad,g,i





# Сбор элементов с вида
custom,category = collect_elements_on_view(doc,
    exclude_categories=[BuiltInCategory.OST_SWallRectOpening,
                        BuiltInCategory.OST_Cameras],
    exclude_classes=[],
    preview="off")
output.freeze()
output.print_md("___")

# for el in custom:
#     level = get_level(el)


bad,g,i = changes(custom)
# if bad:
#     output.print_md("ПРОБЛЕМЫ:")
#     for i,b in enumerate(bad):
#         print(bad)

output.print_md("Всего элементов {}".format(len(custom)))
if bad:
    output.print_md("У {} Были ошибки ".format(len(bad)))
output.print_md("У {} Было задано ".format(g))
output.print_md("У {} Уже заполнено".format(i))
output.print_md("___")
output.unfreeze()