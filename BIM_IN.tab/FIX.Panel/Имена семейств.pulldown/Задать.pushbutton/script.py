# -*- coding: utf-8 -*-

import os
import csv
import codecs
from pyrevit import forms
from Autodesk.Revit.DB import Family, FilteredElementCollector as FEC, Transaction



uidoc = __revit__.ActiveUIDocument  # type: ignore
doc = __revit__.ActiveUIDocument.Document # type: ignore

# 1) Выбрать CSV-файл
csv_path = forms.pick_file(title='Выберите CSV с ID и именами семейств', file_ext='csv')
if not csv_path:
    forms.alert("Файл не выбран", title="Отмена", warn_icon=True)
    raise SystemExit

# 2) Прочитать CSV: ожидаем формат строк: <id>,<new_name>,<category...>
id_to_name = {}
with codecs.open(csv_path, mode='r', encoding='utf-8') as f:
    reader = csv.reader(f)
    for row in reader:
        if not row or len(row) < 2:
            continue
        id_raw = row[0].strip()
        new_name = row[1].strip()
        # пропустим заголовок или мусор
        if not id_raw.isdigit():
            continue
        fam_id_int = int(id_raw)
        if new_name:
            id_to_name[fam_id_int] = new_name

if not id_to_name:
    forms.alert("В CSV не найдено валидных пар ID,Имя.", title="Пустой CSV", warn_icon=True)
    raise SystemExit

# 3) Соберём Family по Id
fams = FEC(doc).OfClass(Family).ToElements()
fam_by_id = {f.Id.IntegerValue: f for f in fams}

# 4) Переименовать. Для семейств корректно использовать Document.RenameElement
renamed = []
skipped_missing = []
skipped_same = []
failed = []

with Transaction(doc, "Переименование семейств из CSV") as t:
    t.Start()
    for fid, new_name in id_to_name.items():
        fam = fam_by_id.get(fid)
        if fam is None:
            skipped_missing.append((fid, new_name, "Family с таким ID не найден в документе"))
            continue

        current = fam.Name
        if current == new_name:
            skipped_same.append((fid, new_name))
            continue
        try:
            fam.Name = new_name
            renamed.append((fid, current, new_name))
        except:
            continue

    t.Commit()

# 5) Краткий отчёт в консоль/скриптовое окно
print("=== Результат переименования семейств ===")
if renamed:
    print("Переименованы:")
    for fid, old, new in renamed:
        print("  ID {}: '{}' -> '{}'".format(fid, old, new))
if skipped_same:
    print("\nПропущены (имя совпадает):")
    for fid, nm in skipped_same:
        print("  ID {}: '{}'".format(fid, nm))
if skipped_missing:
    print("\nПропущены (не найдены по ID):")
    for fid, nm, why in skipped_missing:
        print("  ID {}: '{}' — {}".format(fid, nm, why))
if failed:
    print("\nОшибки:")
    for fid, old, new, err in failed:
        print("  ID {}: '{}' -> '{}': {}".format(fid, old, new, err))
print("=========================================")