# -*- coding: utf-8 -*-
"""
Скрипт для pyRevit (IronPython 2.7)
Задача:
 - Найти общие параметры в проекте Revit
 - Сравнить их с параметрами из подключенного ФОП
 - Проверка идет по GUID и имени
 - Если хотя бы один из параметров отличается — вывести его и вернуть список
"""

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    SharedParameterElement,
    DefinitionFile
)
from Autodesk.Revit.ApplicationServices import Application
from Autodesk.Revit.DB import ExternalDefinition

# Доступ к API Revit
doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application

def get_shared_parameters_from_project(doc):
    """
    Получение всех общих параметров из текущего проекта Revit
    Возвращает список словарей: [{'name': ..., 'guid': ...}, ...]
    """
    shared_params = []
    collector = FilteredElementCollector(doc).OfClass(SharedParameterElement)

    for param_elem in collector:
        definition = param_elem.GetDefinition()
        guid = param_elem.GuidValue  # GUID параметра
        shared_params.append({
            'name': definition.Name,
            'guid': str(guid)
        })
    return shared_params


def get_shared_parameters_from_fop(app):
    """
    Получение всех параметров из подключенного ФОП (файл общих параметров)
    Возвращает список словарей: [{'name': ..., 'guid': ...}, ...]
    """
    fop_path = app.SharedParametersFilename
    if not fop_path:
        raise Exception(u"ФОП не подключен в настройках Revit.")

    def_file = app.OpenSharedParameterFile()
    if not def_file:
        raise Exception(u"Не удалось открыть ФОП: {}".format(fop_path))

    fop_params = []

    # Проходим по всем группам и параметрам в ФОП
    for group in def_file.Groups:
        for definition in group.Definitions:
            if isinstance(definition, ExternalDefinition):
                guid = definition.GUID
                fop_params.append({
                    'name': definition.Name,
                    'guid': str(guid)
                })
    return fop_params


def compare_parameters(project_params, fop_params):
    """
    Сравнение параметров по GUID и имени
    Если хотя бы одно из двух не совпадает — параметр добавляется в список ошибок
    """
    mismatches = []

    # Преобразуем список ФОП в словарь для быстрого поиска по GUID
    fop_dict = {p['guid']: p['name'] for p in fop_params}

    for p_param in project_params:
        guid = p_param['guid']
        name = p_param['name']

        if guid not in fop_dict:
            mismatches.append({
                'type': 'missing_in_fop',
                'guid': guid,
                'name': name
            })
        elif fop_dict[guid] != name:
            mismatches.append({
                'type': 'name_mismatch',
                'guid': guid,
                'name_in_project': name,
                'name_in_fop': fop_dict[guid]
            })

    return mismatches


def main():
    # Получаем параметры из проекта и из ФОП
    project_params = get_shared_parameters_from_project(doc)
    fop_params = get_shared_parameters_from_fop(app)

    # Сравниваем
    mismatches = compare_parameters(project_params, fop_params)

    # Выводим результаты
    if not mismatches:
        print(u"✅ Все параметры совпадают по имени и GUID")
    else:
        print(u"⚠ Найдены несоответствия:")
        for m in mismatches:
            # if m['type'] == 'missing_in_fop':
            #     print(u"- Параметр '{}' (GUID: {}) отсутствует в ФОП".format(m['name'], m['guid']))
            if m['type'] == 'name_mismatch':
                print(u"- GUID {}: имя в проекте '{}', а в ФОП '{}'".format(
                    m['guid'], m['name_in_project'], m['name_in_fop']
                ))
    return mismatches


if __name__ == "__main__":
    result = main()