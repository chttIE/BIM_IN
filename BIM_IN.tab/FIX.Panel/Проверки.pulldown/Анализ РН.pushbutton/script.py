# -*- coding: utf-8 -*-
__title__ = "Анализ РН"
__author__ = 'NistratovIlia'
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import Workset, WorksetTable, WorksetKind
from System.Collections.Generic import List

doc = __revit__.ActiveUIDocument.Document

def get_empty_worksets(doc):
    empty_worksets = []
    worksets = FilteredWorksetCollector(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()

    for ws in worksets:
        ws_id = ws.Id
        # Получаем все элементы в рабочем наборе
        collector = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
        elements_in_ws = [el for el in collector if el.WorksetId == ws_id]

        if not elements_in_ws:
            empty_worksets.append(ws)

    return empty_worksets


# Выводим список пустых РН
empty_ws = get_empty_worksets(doc)
if empty_ws:
    print("🔍 Найдены пустые рабочие наборы:")
    for ws in empty_ws:
        print("🗂️  {} (ID: {})".format(ws.Name, ws.Id.IntegerValue))
else:
    print("✅ Все рабочие наборы содержат элементы.")
