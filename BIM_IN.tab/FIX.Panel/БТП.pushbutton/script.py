# -*- coding: utf-8 -*-
# Release v1 11/02/23
__title__ = 'БТП'
__doc__ = ''
__author__ = 'IliaNistratov'


from Autodesk.Revit.DB import *
from pyrevit import revit, script
import os

doc = revit.doc
output = script.get_output()
lfy = output.linkify
# Таблица координат (в мм)
coordinate_table = {
    "K28":    (12141150.5, 9451626.3),
    "K29":    (12014289.4, 9390434.2),
    "K30":    (11989821.4, 9502908.5),
    "K31":    (12017323.8, 9559253.8),
    "K29-31": (12014289.4, 9390434.2)
}

# Получаем имя модели (без пути и расширения)
model_name = os.path.splitext(os.path.basename(doc.PathName))[0]

# Ищем ключ в имени модели
found_key = None
for key in coordinate_table:
    if key in model_name:
        found_key = key
        break

if not found_key:
    output.print_md("❌ Не удалось определить код модели по имени файла: `{}`".format(model_name))
else:
    east_mm, north_mm = coordinate_table[found_key]
    east_ft = east_mm / 304.8
    north_ft = north_mm / 304.8

    output.print_md("### ➕ Установка координат по коду **{}**".format(found_key))
    output.print_md("- Восток/Запад (мм): `{}`".format(east_mm))
    output.print_md("- Север/Юг (мм): `{}`".format(north_mm))

    # Получаем базовую точку проекта
    base_point = BasePoint.GetProjectBasePoint(doc)
    output.print_md(lfy(base_point.Id))
    with Transaction(doc, "Установка координат БТП") as t:
        t.Start()

        # Открепляем точку (Shared = False)

 

        # Устанавливаем координаты
        base_point.get_Parameter(BuiltInParameter.BASEPOINT_EASTWEST_PARAM).Set(east_ft)
        base_point.get_Parameter(BuiltInParameter.BASEPOINT_NORTHSOUTH_PARAM).Set(north_ft)

        # Закрепляем точку обратно


        t.Commit()

    output.print_md("✅ Координаты успешно установлены для `{}`".format(model_name))
