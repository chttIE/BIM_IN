# -*- coding: utf-8 -*-
__title__ = 'Добавить параметр\nADSK_Комплект'
__doc__ = 'Добавляет общий параметр ADSK_Комплект для листов, видов и спецификаций по экземпляру.'
__author__ = 'IliaNistratov'

from Autodesk.Revit.DB import *
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.UI import TaskDialog
from pyrevit import script
import os

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
output = script.get_output()
output.set_title(__title__)
output.set_width(600)

# Категории, к которым добавим параметр
target_categories = [
    BuiltInCategory.OST_Sheets,         # Листы
    BuiltInCategory.OST_Views,          # Виды
    BuiltInCategory.OST_Schedules       # Спецификации
]

param_name = "ADSK_Комплект"
group_name = "05 Необязательные ОБЩИЕ"

def get_shared_param_def(app, group_name, param_name):
    # Получает определение общего параметра из активного FOP
    file = app.OpenSharedParameterFile()
    if not file:
        TaskDialog.Show("Ошибка", "Не найден файл общих параметров (FOP).")
        return None

    for group in file.Groups:
        if group.Name == group_name:
            for definition in group.Definitions:
                if definition.Name == param_name:
                    return definition
    return None

def add_parameter_to_project(doc, definition, target_bics):
    category_set = CategorySet()
    for bic in target_bics:
        cat = doc.Settings.Categories.get_Item(bic)
        if cat:
            category_set.Insert(cat)

    binding = InstanceBinding(category_set)
    param_bindings = doc.ParameterBindings

    if param_bindings.Insert(definition, binding, BuiltInParameterGroup.PG_DATA):
        output.print_md("✅ Параметр **{}** добавлен в проект.".format(param_name))
    else:
        output.print_md("⚠️ Параметр **{}** уже существует или не удалось добавить.".format(param_name))


with Transaction(doc, "Добавление общего параметра ADSK_Комплект") as t:
    t.Start()

    definition = get_shared_param_def(app, group_name, param_name)

    if definition:
        add_parameter_to_project(doc, definition, target_categories)
    else:
        output.print_md("❌ Параметр **{}** не найден в группе **{}**.".format(param_name, group_name))

    t.Commit()