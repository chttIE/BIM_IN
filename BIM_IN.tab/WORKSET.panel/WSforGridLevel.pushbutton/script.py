# -*- coding: utf-8 -*
from pyrevit import forms, script
from Autodesk.Revit.DB import *

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
output.set_width(600)

# Функция переноса и закрепления
def move_and_pin(elements, workset_id, label):
    count = 0
    for el in elements:
        param = el.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM)
        if param and param.StorageType == StorageType.Integer:
            param.Set(workset_id.IntegerValue)
            el.Pinned = True
            count += 1
    output.print_md("- ✅ Перенесено {} элементов типа **{}** в рабочий набор **{}**.".format(count, label, target_ws_name))

target_ws = forms.SelectFromList.show(FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets(),
                                    multiselect = False,
                                    title='Выбор РН для осей и уровней',                                                    
                                    name_attr='Name',
                                    button_name='Выбрать')
if not target_ws:
    script.exit()

target_ws_id = target_ws.Id
target_ws_name = target_ws.Name

# Сбор элементов
grids = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Grids).WhereElementIsNotElementType().ToElements()
levels = FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements()


# Транзакция
with Transaction(doc, "Перенос в {}".format(target_ws_name)) as t:
    t.Start()
    move_and_pin(grids, target_ws_id, "Оси")
    move_and_pin(levels, target_ws_id, "Уровни")
    t.Commit()

output.print_md("### 🟢 Готово. Все оси и уровни перенесены и закреплены.")