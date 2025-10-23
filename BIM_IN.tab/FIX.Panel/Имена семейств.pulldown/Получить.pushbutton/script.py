# -*- coding: utf-8 -*-

import os
import csv
import codecs
from pyrevit import forms
from Autodesk.Revit.DB import Family, FilteredElementCollector as FEC

doc = __revit__.ActiveUIDocument.Document  # type: ignore

# Получение всех семейств
fams = FEC(doc).OfClass(Family).WhereElementIsNotElementType().ToElements()
fam_data = {}

for f in fams:
    f_name = f.Name
    if "ADSK" in f_name:
        continue  # Пропуск семейств Autodesk
    f_Cat = f.FamilyCategory.Name if f.FamilyCategory else "Нет категории"
    f_id = f.Id.IntegerValue
    fam_data[f_id] = [f_name, f_Cat]
    print("{} | {} | {}".format(f_id, f_name, f_Cat))

# Выбор папки
path = forms.pick_folder(title='Выберите папку для сохранения CSV файла')
if not path:
    forms.alert("Папка не выбрана. Операция отменена.", title="Отмена", warn_icon=True)
    raise Exception("Операция отменена пользователем")

# Имя файла
name_csv = doc.Title.replace('.rvt', '') + '_family_data.csv'
full_path = os.path.join(path, name_csv)

# Запись CSV с BOM (UTF-8)
with codecs.open(full_path, mode='w', encoding='utf-8') as file:
    # Запись BOM вручную для Excel
    file.write(codecs.BOM_UTF8.decode('utf-8'))
    
    writer = csv.writer(file)
    writer.writerow(['Family ID', 'Family Name', 'Category'])
    for f_id, (f_name, f_Cat) in fam_data.items():
        writer.writerow([f_id, f_name, f_Cat])

print("CSV файл сохранен: {}".format(full_path))
