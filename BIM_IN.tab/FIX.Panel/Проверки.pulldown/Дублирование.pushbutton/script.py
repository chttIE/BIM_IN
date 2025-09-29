#coding: utf-8
from collections import defaultdict, deque
from Autodesk.Revit.DB import *
from pyrevit import script,output,revit

title = "Проверка дублирования"
output = script.get_output()
output.close_others(all_open_outputs=True)
output.set_title(title)
output.set_width(900)

lfy = output.linkify
doc = revit.doc
user = __revit__.Application.Username



def find_duplicates_groups(d):
    # Получаем все предупреждения
    all_warnings = d.GetWarnings()
    if not all_warnings:
        script.exit()

    # Фильтруем только предупреждения о дублировании
    duplicates_warnings = []
    for w in all_warnings:
        if "дублированию" in w.GetDescriptionText():
            duplicates_warnings.append(w)
    if not duplicates_warnings:
        script.exit()

    # Создаём граф: элемент -> связанные элементы
    graph = defaultdict(set)
    for w in duplicates_warnings:
        elements = w.GetFailingElements()
        for elem1 in elements:
            for elem2 in elements:
                if elem1 != elem2:
                    graph[elem1].add(elem2)
                    graph[elem2].add(elem1)

    # Найти все компоненты связности
    visited = set()
    groups = []

    def bfs(start):
        queue = deque([start])
        component = set()
        while queue:
            node = queue.popleft()
            if node not in visited:
                visited.add(node)
                component.add(node)
                queue.extend(graph[node])
        return component

    for node in graph:
        if node not in visited:
            groups.append(bfs(node))

    # Определяем оригиналы и копии
    result = []
    for group in groups:
        sorted_ids = sorted(group, key=lambda x: x.IntegerValue)  # Сортируем по ID
        original = sorted_ids[0]  # Минимальный ID - оригинал
        copies = sorted_ids[1:]  # Остальные - копии
        result.append({"original": original, "copies": copies})

    return result

def select_all(element_ids, MAX_ELEMENTS_PER_LINK=100, title="Выделить"):
    from Autodesk.Revit.DB import ElementId

    # Проверка, что все элементы являются ElementId
    if not all(isinstance(eid, ElementId) for eid in element_ids):
        raise ValueError("Список element_ids содержит элементы, которые не являются ElementId.")
    
    # Если элементов больше ограничения, создаём несколько ссылок
    total_elements = len(element_ids)
    if total_elements > MAX_ELEMENTS_PER_LINK:
        chunks = [element_ids[i:i + MAX_ELEMENTS_PER_LINK] for i in range(0, total_elements, MAX_ELEMENTS_PER_LINK)]
        for idx, chunk in enumerate(chunks, start=1):
            link_text = "{} часть ({}/{})".format(title, idx, len(chunks))
            link = lfy(chunk, link_text)
            output.print_md(">{}".format(link))
    else:
        # Если элементов меньше ограничения, создаём одну ссылку
        link_text = "{} ({})".format(title, total_elements)
        link = lfy(element_ids, link_text)
        output.print_md(">{}".format(link))
# Пример вызова функции:

groups = find_duplicates_groups(doc)
output.print_md("___")
output.print_md("Модель: **{}**".format(doc.Title))
output.print_md("___")
wrngs_data=[]
copies_lst = []
for group in groups:
    wrngs_data.append(( lfy(group["original"],"Выбрать"),"**{} **".format(len(group["copies"])) + lfy(group["copies"],"Выбрать все")))  
    copies_lst.append(group["copies"])

wrngs_data = [(i + 1,) + x for i, x in enumerate(wrngs_data)]
output.print_table(
    table_data=wrngs_data,
    title="**Дублирований {}.**".format(wrngs_data.Count),
    columns=["№", "Оригинал", "Дубликаты"],
    formats=[ '', '', ''])

elids = [eid for sublist in copies_lst for eid in sublist]
select_all(elids,MAX_ELEMENTS_PER_LINK=100,title="Выделить")
