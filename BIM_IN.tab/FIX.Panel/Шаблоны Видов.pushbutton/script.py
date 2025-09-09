# -*- coding: utf-8 -*-
__title__ = 'Переименовать\nшаблоны'
__doc__ = 'Переименовывает шаблоны видов по BIM-стандарту ADSK.'
__author__ = 'IliaNistratov'

from pyrevit import script
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
output.set_title(__title__)
output.set_width(600)

def normalize_name(name):
    """Заменяет пробелы на подчёркивания, убирает двойные подчёркивания, приводит к верхнему регистру."""
    name = name.replace(" ", "_")
    while "__" in name:
        name = name.replace("__", "_")
    while "!" in name:
        name = name.replace("!", "")
        
    return name.upper()

def rename_view_templates_with_prefix(doc):
    """
    Переименовывает шаблоны видов с добавлением префиксов:
    - О_П_ для планов
    - О_Р_ для разрезов
    - О_Ф_ для фасадов
    - О_С_ для спецификаций
    """
    templates = FilteredElementCollector(doc).OfClass(View).ToElements()
    renamed = []

    for view in templates:
        if not view.IsTemplate:
            continue

        original_name = view.Name
        if "ADSK_" in original_name:
            continue
        view_type = view.ViewType

        # Префикс по типу вида
        if view_type in [ViewType.FloorPlan, ViewType.CeilingPlan]:
            prefix = "О_П_"
        elif view_type == ViewType.Section:
            prefix = "О_Р_"
        elif view_type == ViewType.Elevation:
            prefix = "О_Ф_"
        elif view_type == ViewType.Schedule:
            prefix = "О_С_"
        else:
            continue  # Пропустить неподдерживаемые типы

        # Проверяем, есть ли уже нужный префикс
        if not original_name.startswith(prefix):
            name_with_prefix = "{}{}".format(prefix, original_name)
        else:
            name_with_prefix = original_name

        # Применяем нормализацию в самом конце
        final_name = normalize_name(name_with_prefix)

        # Переименование, если имя изменилось
        if final_name != original_name:
            try:
                view.Name = final_name
                renamed.append((original_name, final_name))
            except Exception as e:
                print("❌ Не удалось переименовать {}: {}".format(original_name, e))

    # Вывод результатов
    print("✅ Переименовано шаблонов: {}".format(len(renamed)))
    for old, new in renamed:
        print("- {} → {}".format(old, new))


with Transaction(doc, "Переименование шаблонов видов") as t:
    t.Start()
    rename_view_templates_with_prefix(doc)
    t.Commit()
