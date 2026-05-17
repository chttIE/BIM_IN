# -*- coding: utf-8 -*-
__title__ = "Заполнить\nпараметры"
__author__ = "IliaNistratov"
__doc__ = "Заполняет RUS_MSSK_Element_Name по коду RUS_MSSK_Element_Code из классификатора МССК."

import codecs
import os
import re

from Autodesk.Revit.DB import FilteredElementCollector, StorageType, Transaction
from pyrevit import forms, script


doc = __revit__.ActiveUIDocument.Document  # type: ignore
uidoc = __revit__.ActiveUIDocument  # type: ignore
output = script.get_output()
output.close_others(all_open_outputs=True)


CODE_PARAMETER = "RUS_MSSK_Element_Code"
NAME_PARAMETER = "RUS_MSSK_Element_Name"
CLASSIFIER_FILE = u"Классификатор МССК_v5.txt"


def get_script_folder():
    try:
        return os.path.dirname(__file__)
    except:
        return os.getcwd()


def get_classifier_path():
    return os.path.join(get_script_folder(), CLASSIFIER_FILE)


def clean_text(value):
    if value is None:
        return ""
    return value.strip()


def normalize_code(value):
    value = clean_text(value)
    value = value.replace(u"\ufeff", "")
    value = value.replace("\x00", "")
    value = value.replace(u"\xa0", " ")
    value = value.replace("\t", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip().upper()


def get_type_element(element):
    try:
        type_id = element.GetTypeId()
        if type_id and type_id.IntegerValue > 0:
            return doc.GetElement(type_id)
    except:
        pass

    return None


def lookup_parameter_instance_or_type(element, parameter_name):
    parameter = element.LookupParameter(parameter_name)
    if parameter:
        return parameter, element, "instance"

    type_element = get_type_element(element)
    if type_element:
        parameter = type_element.LookupParameter(parameter_name)
        if parameter:
            return parameter, type_element, "type"

    return None, None, None


def get_parameter_text(element, parameter_name):
    parameter, owner, scope = lookup_parameter_instance_or_type(element, parameter_name)
    if not parameter:
        return None

    if parameter.StorageType == StorageType.String:
        return clean_text(parameter.AsString())

    return clean_text(parameter.AsValueString())


def set_parameter_text(element, parameter_name, value):
    parameter, owner, scope = lookup_parameter_instance_or_type(element, parameter_name)
    if not parameter:
        return False, "missing parameter", scope, owner

    return set_existing_string_parameter(parameter, value) + (scope, owner)


def set_existing_string_parameter(parameter, value):
    if parameter.IsReadOnly:
        return False, "readonly parameter"

    if parameter.StorageType != StorageType.String:
        return False, "not a string parameter"

    current_value = clean_text(parameter.AsString())
    new_value = clean_text(value)

    if current_value == new_value:
        return False, "same value"

    parameter.Set(new_value)
    return True, None


def get_element_key(element):
    try:
        return element.Id.IntegerValue
    except:
        return id(element)


def parse_classifier_line(line):
    line = clean_text(line)
    if not line:
        return None

    parts = [clean_text(part) for part in line.split("\t")]
    parts = [part for part in parts if part]

    if len(parts) < 2:
        return None

    code = normalize_code(parts[0])
    name = parts[1]

    if not code or not name:
        return None

    return code, name


def read_classifier_lines(path):
    with open(path, "rb") as classifier_file:
        raw_data = classifier_file.read()

    if raw_data.startswith(codecs.BOM_UTF16_LE) or raw_data.startswith(codecs.BOM_UTF16_BE):
        return raw_data.decode("utf-16").splitlines(), "utf-16"

    if raw_data.startswith(codecs.BOM_UTF8):
        return raw_data.decode("utf-8-sig").splitlines(), "utf-8-sig"

    encodings = ["utf-8-sig", "cp1251", "utf-16"]
    last_error = None

    for encoding in encodings:
        try:
            return raw_data.decode(encoding).splitlines(), encoding
        except Exception as error:
            last_error = error

    raise last_error


def read_classifier(path):
    code_to_name = {}
    lines, encoding = read_classifier_lines(path)
    output.print_md("- Classifier encoding: **{}**".format(encoding))

    for line_number, line in enumerate(lines, 1):
        parsed = parse_classifier_line(line)
        if not parsed:
            continue

        code, name = parsed
        if code not in code_to_name:
            code_to_name[code] = name
        else:
            output.print_md(
                "- Duplicate code in classifier ignored: **{}**, line **{}**".format(
                    code,
                    line_number,
                )
            )

    return code_to_name


def filter_elements_with_code_parameter(elements):
    result = []

    for element in elements:
        if not element:
            continue

        parameter, owner, scope = lookup_parameter_instance_or_type(element, CODE_PARAMETER)
        if parameter:
            result.append(element)

    return result


def collect_selected_elements_with_code_parameter():
    selected_ids = list(uidoc.Selection.GetElementIds())
    if not selected_ids:
        return [], False

    elements = [doc.GetElement(element_id) for element_id in selected_ids]
    return filter_elements_with_code_parameter(elements), True


def collect_project_elements_with_code_parameter():
    elements = FilteredElementCollector(doc).WhereElementIsNotElementType().ToElements()
    return filter_elements_with_code_parameter(elements)


def collect_elements_with_code_parameter():
    selected_elements, has_selection = collect_selected_elements_with_code_parameter()
    if has_selection:
        return selected_elements, "selection"

    return collect_project_elements_with_code_parameter(), "project"


def fill_names(elements, classifier):
    changed = 0
    already_filled = 0
    skipped_type_duplicates = 0
    empty_code = []
    code_not_found = []
    write_errors = []
    processed_type_names = set()

    with Transaction(doc, "BIM_IN | RUS_MSSK_Element_Name") as transaction:
        transaction.Start()

        for element in elements:
            name_parameter, name_owner, name_scope = lookup_parameter_instance_or_type(element, NAME_PARAMETER)
            if not name_parameter:
                write_errors.append((element, "missing parameter"))
                continue

            if name_scope == "type":
                type_key = get_element_key(name_owner)
                if type_key in processed_type_names:
                    skipped_type_duplicates += 1
                    continue

            code = normalize_code(get_parameter_text(element, CODE_PARAMETER))
            if not code:
                empty_code.append(element)
                continue

            name = classifier.get(code)
            if not name:
                code_not_found.append((element, code))
                continue

            ok, reason = set_existing_string_parameter(name_parameter, name)
            if ok:
                changed += 1
            elif reason == "same value":
                already_filled += 1
            else:
                write_errors.append((element, reason))

            if name_scope == "type":
                processed_type_names.add(type_key)

        transaction.Commit()

    return {
        "changed": changed,
        "already_filled": already_filled,
        "skipped_type_duplicates": skipped_type_duplicates,
        "empty_code": empty_code,
        "code_not_found": code_not_found,
        "write_errors": write_errors,
    }


def print_report(elements_count, classifier_count, result, collect_mode):
    output.print_md("## Заполнение RUS_MSSK_Element_Name")
    if collect_mode == "selection":
        output.print_md("- Режим: **предварительный выбор**")
    else:
        output.print_md("- Режим: **все элементы проекта**")
    output.print_md("- Элементов с параметром `{}`: **{}**".format(CODE_PARAMETER, elements_count))
    output.print_md("- Кодов в классификаторе: **{}**".format(classifier_count))
    output.print_md("- Записано значений: **{}**".format(result["changed"]))
    output.print_md("- Уже было заполнено: **{}**".format(result["already_filled"]))
    output.print_md("- Пропущено повторов типоразмера: **{}**".format(result["skipped_type_duplicates"]))
    output.print_md("- Пустой код: **{}**".format(len(result["empty_code"])))
    output.print_md("- Код не найден: **{}**".format(len(result["code_not_found"])))
    output.print_md("- Ошибок записи: **{}**".format(len(result["write_errors"])))

    if result["code_not_found"]:
        output.print_md("### Коды не найдены")
        for element, code in result["code_not_found"][:100]:
            output.print_md("- {} | `{}`".format(output.linkify(element.Id), code))
        if len(result["code_not_found"]) > 100:
            output.print_md("- ...и еще **{}**".format(len(result["code_not_found"]) - 100))

    if result["write_errors"]:
        output.print_md("### Ошибки записи")
        for element, reason in result["write_errors"][:100]:
            output.print_md("- {} | {}".format(output.linkify(element.Id), reason))
        if len(result["write_errors"]) > 100:
            output.print_md("- ...и еще **{}**".format(len(result["write_errors"]) - 100))


classifier_path = get_classifier_path()

if not os.path.exists(classifier_path):
    forms.alert(
        "Не найден файл классификатора:\n\n{}\n\n"
        "Положи файл рядом со script.py этой кнопки и запусти инструмент снова.".format(classifier_path),
        title="Классификатор МССК",
    )
    script.exit()

classifier = read_classifier(classifier_path)
if not classifier:
    forms.alert(
        "Классификатор пустой или не удалось прочитать строки формата:\n"
        "ЭЛ 30 16 83 14<TAB>Прокладка<TAB>4",
        title="Классификатор МССК",
    )
    script.exit()

elements_with_code, collect_mode = collect_elements_with_code_parameter()
if not elements_with_code:
    forms.alert(
        "Не найдено элементов с параметром {}.\n\n"
        "Если был предварительный выбор, проверь, что выбранные элементы имеют этот параметр у экземпляра или типоразмера.".format(CODE_PARAMETER),
        title="Заполнение параметров",
    )
    script.exit()

result = fill_names(elements_with_code, classifier)
print_report(len(elements_with_code), len(classifier), result, collect_mode)
