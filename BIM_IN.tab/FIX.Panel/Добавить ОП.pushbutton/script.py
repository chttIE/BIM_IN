# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import *
from pyrevit import  DB, forms, script
output = script.get_output()
output.close_others(False)
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document #type:Document
# 1. Открываем файл общих параметров
sp_file = app.OpenSharedParameterFile()
if not sp_file:
    forms.alert("Не найден файл общих параметров", exitscript=True)

# 2. Собираем все параметры в словарь
param_dict = {}
for group in sp_file.Groups:
    for definition in group.Definitions:
        param_dict["[{}]_{}".format(group.Name, definition.Name)] = definition

# 3. Выбираем параметр
param_name = forms.SelectFromList.show(sorted(param_dict.keys()), 
                                       title="Выберите общий параметр", 
                                       multiselect=True)
if not param_name:
    forms.alert("Параметр не выбран", exitscript=True)

selected_definition = [param_dict[param_name] for param_name in param_name]

# 4. Выбираем категории
categories = doc.Settings.Categories
cat_names = sorted([c.Name for c in categories])
selected_cats = forms.SelectFromList.show(cat_names, 
                                          title="Выберите категории", 
                                          multiselect=True)

cat_set = DB.CategorySet()
for name in selected_cats:
    cat_set.Insert(categories.get_Item(name))

# 5. Выбираем тип/экземпляр
param_type = forms.alert("Сделать параметр типовым?", options=["Да", "Нет"])
if param_type == "Да":  binding = DB.TypeBinding(cat_set)
else: binding = DB.InstanceBinding(cat_set)

# 6. Выбираем группу параметров
group_options = list(DB.BuiltInParameterGroup.GetValues(DB.BuiltInParameterGroup))

# Делаем словарь {русское имя : enum}
group_dict = {}
for g in group_options:
    try:
        name_ru = DB.LabelUtils.GetLabelFor(g)
    except:
        # На случай, если у группы нет локализации
        name_ru = str(g)
    group_dict[name_ru] = g

# Выбор по русским именам
selected_group_name = forms.SelectFromList.show(sorted(group_dict.keys()), title="Выберите группу параметров")
if not selected_group_name:
    script.exit()

selected_group_enum = group_dict[selected_group_name]
# 7. Добавляем параметр в проект
with DB.Transaction(doc, "Добавление общего параметра") as t:
    t.Start()
    for p in selected_definition:
        try:
            doc.ParameterBindings.Insert(p,binding,selected_group_enum)
            output.print_md("Параметр '{}' успешно добавлен!".format(p.Name))
        except Exception as e :
            output.print_md("Параметр '{}' НЕ добавлен! Ошибка {}".format(p.Name,str(e)))
    t.Commit()
