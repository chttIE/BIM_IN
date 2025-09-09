# -*- coding: utf-8 -*-
__title__ = "Удалить неисп.\nпараметры"
__author__ = "IliaNistratov"
__doc__ = "Поиск и удаление проектных параметров, не применяемых в модели"

from Autodesk.Revit.DB import *
from pyrevit import script, forms, revit

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
lfy = output.linkify
output.close_others(all_open_outputs=True)
# Получаем все проектные параметры
binding_map = doc.ParameterBindings
it = binding_map.ForwardIterator()
it.Reset()

unused_params = {}
param_def_map = {}
shared_param_names = set()

while it.MoveNext():
    definition = it.Key
    param_def_map[definition.Name] = definition
    unused_params[definition.Name] = {
        "definition": definition,
        "used_in_elements": False,
        "used_in_filters": False,
        "used_in_schedules": False
    }

# Определяем общие параметры через ParameterElement
all_param_elements = FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()
for pe in all_param_elements:
    definition = pe.GetDefinition()
    if isinstance(pe,SharedParameterElement):
        shared_param_names.add(pe.Name)

# Проверка использования в элементах
collector = FilteredElementCollector(doc).WhereElementIsNotElementType()
for el in collector:
    for param in el.Parameters:
        try:
            name = param.Definition.Name
            if name in unused_params:
                if param.HasValue and (param.AsString() or param.AsValueString() or param.AsInteger() or param.AsDouble()):
                    unused_params[name]["used_in_elements"] = True
        except:
            continue

# Проверка в "Сведения о проекте"
project_info = doc.ProjectInformation
for param in project_info.Parameters:
    try:
        name = param.Definition.Name
        if name in unused_params:
            if param.HasValue and (param.AsString() or param.AsValueString() or param.AsInteger() or param.AsDouble()):
                unused_params[name]["used_in_elements"] = True
    except:
        continue


# Проверка в фильтрах


def get_Filters(filter_elem):
    try:
        return filter_elem.GetElementFilter().GetFilters()
    except:
        return None

def get_rules(flt):
    try:
        return flt.GetRules()
    except: 
        return []

def get_param_rule(rule, doc):
    try:
        param_id = rule.GetRuleParameter()
        if param_id.IntegerValue < 0:
            # Встроенный параметр
            return LabelUtils.GetLabelFor(BuiltInParameter(param_id.IntegerValue))
        else:
            # Пользовательский параметр
            elem = doc.GetElement(param_id)
            return elem.Name if elem else "UNKNOWN"
    except:
        return "ERROR"


# Получение фильтров
filters = FilteredElementCollector(doc).OfClass(ParameterFilterElement).ToElements()

for filter_elem in filters:
    inner_filters = get_Filters(filter_elem)
    if not inner_filters:
        continue

    for flt in inner_filters:
        rules = get_rules(flt)
        for rule in rules:
            pname = get_param_rule(rule, doc)
            if pname in unused_params:
                unused_params[pname]["used_in_filters"] = True

# Проверка в спецификациях
schedules = FilteredElementCollector(doc).OfClass(ViewSchedule).ToElements()
for sched in schedules:
    # print(lfy(sched.Id))
    try:
        definition = sched.Definition
        for field_id in definition.GetFieldOrder():
            field = definition.GetField(field_id)
            pname = field.GetName()
            # print(pname)
            if pname in unused_params:
                unused_params[pname]["used_in_schedules"] = True
    except:
        continue

# Список неиспользуемых
unused_names = [name for name, data in unused_params.items()
                if not data["used_in_elements"] and not data["used_in_filters"] and not data["used_in_schedules"]]

if not unused_names:
    output.print_md("✅ Все параметры используются!")
    script.exit()

# Формирование списка для отображения (добавляем "[ОБЩИЙ]" к общим)
display_names = []
name_map = {}  # отображаемое -> реальное имя

for name in sorted(unused_names):
    display = "[ОБЩИЙ] {}".format(name) if name in shared_param_names else name
    display_names.append(display)
    name_map[display] = name

# Выбор параметров
selected_display_names = forms.SelectFromList.show(
    display_names,
    multiselect=True,
    title='Неиспользуемые параметры',
    width=450,
    height=800,
    button_name='Удалить'
)

if not selected_display_names:
    script.exit()

# Удаление через doc.Delete
error_list = []
deleted = []

with Transaction(doc, "Удаление неисп. параметров") as t:
    t.Start()
    for display_name in selected_display_names:
        pname = name_map[display_name]

        # Заново ищем ParameterElement по имени (т.к. некоторые уже могли быть удалены)
        param_elem = next(
            (pe for pe in FilteredElementCollector(doc).OfClass(ParameterElement).ToElements()
             if pe.Name == pname),
            None
        )

        if param_elem:
            try:
                doc.Delete(param_elem.Id)
                output.print_md('- 🗑 Удален параметр: **{}**'.format(display_name))
                deleted.append(pname)
            except:
                error_list.append('- ❌ Не удалось удалить параметр: **{}**'.format(display_name))
        else:
            error_list.append('- ⚠️ Параметр "{}" не найден в ParameterElement'.format(display_name))
    t.Commit()


# Ошибки
if error_list:
    output.print_md("## ⚠️ Ошибки:")
    for err in error_list:
        output.print_md(err)
