# -*- coding: utf-8 -*-
__title__ = "Задать\nкорпус/секцию"
__author__ = "IliaNistratov"
__doc__ = "Записывает ADSK_Номер корпуса и ADSK_Номер секции для элементов, найденных как в кнопке заполнения параметров МССК."

from Autodesk.Revit.DB import CheckoutStatus, FilteredElementCollector, StorageType, Transaction, WorksharingUtils
from pyrevit import forms, script
from rpw.ui.forms import FlexForm, Label, TextBox, Separator, Button


doc = __revit__.ActiveUIDocument.Document  # type: ignore
uidoc = __revit__.ActiveUIDocument  # type: ignore
output = script.get_output()
output.close_others(all_open_outputs=True)


CODE_PARAMETER = "RUS_MSSK_Element_Code"
BUILDING_PARAMETER = u"ADSK_Номер корпуса"
SECTION_PARAMETER = u"ADSK_Номер секции"

try:
    text_type = unicode
except NameError:
    text_type = str


def clean_text(value):
    if value is None:
        return ""
    return text_type(value).strip()


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


def get_element_key(element):
    try:
        return element.Id.IntegerValue
    except:
        return id(element)


def get_parameter_key(owner, parameter_name):
    return "{}::{}".format(get_element_key(owner), parameter_name)


def get_worksharing_lock_reason(element):
    if not doc.IsWorkshared:
        return None

    try:
        checkout_status = WorksharingUtils.GetCheckoutStatus(doc, element.Id)
    except Exception as error:
        return u"ошибка проверки занятости элемента: {}".format(error)

    if checkout_status != CheckoutStatus.OwnedByOtherUser:
        return None

    owner = None
    try:
        tooltip_info = WorksharingUtils.GetWorksharingTooltipInfo(doc, element.Id)
        owner = clean_text(tooltip_info.Owner)
    except:
        pass

    if owner:
        return u"элемент занят другим пользователем: {}".format(owner)

    return u"элемент занят другим пользователем"


def get_parameter_value_text(parameter):
    if parameter.StorageType == StorageType.String:
        return clean_text(parameter.AsString())

    return clean_text(parameter.AsValueString())


def parse_integer(value):
    value = clean_text(value)
    try:
        return int(value)
    except:
        raise ValueError(u"значение должно быть целым числом")


def parse_double(value):
    value = clean_text(value).replace(",", ".")
    try:
        return float(value)
    except:
        raise ValueError(u"значение должно быть числом")


def set_parameter_value(parameter, value):
    if parameter.IsReadOnly:
        return False, "readonly parameter"

    value = clean_text(value)
    current_value = get_parameter_value_text(parameter)
    if current_value == value:
        return False, "same value"

    try:
        if parameter.StorageType == StorageType.String:
            parameter.Set(value)
        elif parameter.StorageType == StorageType.Integer:
            parameter.Set(parse_integer(value))
        elif parameter.StorageType == StorageType.Double:
            parameter.Set(parse_double(value))
        else:
            return False, "unsupported parameter type"
    except Exception as error:
        return False, text_type(error)

    return True, None


def show_values_form():
    components = [
        Label(u"Значение для номера корпуса".format(BUILDING_PARAMETER)),
        TextBox("building_number", Text="", Height=28),
        Label(u"Значение для номера секции".format(SECTION_PARAMETER)),
        TextBox("section_number", Text="", Height=28),
        Separator(),
        Button(u"Записать", Height=30),
    ]

    form = FlexForm(u"Корпус и секция", components)
    form.ShowDialog()

    if not form.values:
        return None

    building_number = clean_text(form.values.get("building_number"))
    section_number = clean_text(form.values.get("section_number"))
    if not building_number or not section_number:
        forms.alert(
            u"Нужно заполнить оба значения: корпус и секцию.",
            title=u"Корпус и секция",
        )
        return None

    return {
        BUILDING_PARAMETER: building_number,
        SECTION_PARAMETER: section_number,
    }


def confirm_write(elements_count, values, collect_mode):
    if collect_mode == "selection":
        mode_text = u"предварительный выбор"
    else:
        mode_text = u"все элементы проекта"

    message = (
        u"Будет обработано элементов с параметром {code}: {count}\n"
        u"Режим: {mode}\n\n"
        u"{building}: {building_value}\n"
        u"{section}: {section_value}"
    ).format(
        code=CODE_PARAMETER,
        count=elements_count,
        mode=mode_text,
        building=BUILDING_PARAMETER,
        building_value=values[BUILDING_PARAMETER],
        section=SECTION_PARAMETER,
        section_value=values[SECTION_PARAMETER],
    )

    result = forms.alert(
        message,
        title=u"Корпус и секция",
        options=[u"Записать", u"Отмена"],
    )
    return result == u"Записать"


def fill_parameters(elements, values):
    changed = 0
    already_filled = 0
    skipped_type_duplicates = 0
    missing_parameters = []
    write_errors = []
    processed_parameter_keys = set()

    with Transaction(doc, u"BIM_IN | ADSK_Номер корпуса и секции") as transaction:
        transaction.Start()

        for element in elements:
            for parameter_name, value in values.items():
                parameter, owner, scope = lookup_parameter_instance_or_type(element, parameter_name)
                if not parameter:
                    missing_parameters.append((element, parameter_name))
                    continue

                parameter_key = get_parameter_key(owner, parameter_name)
                if scope == "type" and parameter_key in processed_parameter_keys:
                    skipped_type_duplicates += 1
                    continue

                lock_reason = get_worksharing_lock_reason(owner)
                if lock_reason:
                    write_errors.append((owner, parameter_name, lock_reason))
                    processed_parameter_keys.add(parameter_key)
                    continue

                is_changed, issue = set_parameter_value(parameter, value)
                if is_changed:
                    changed += 1
                elif issue == "same value":
                    already_filled += 1
                else:
                    write_errors.append((owner, parameter_name, issue))

                if scope == "type":
                    processed_parameter_keys.add(parameter_key)

        transaction.Commit()

    return {
        "changed": changed,
        "already_filled": already_filled,
        "skipped_type_duplicates": skipped_type_duplicates,
        "missing_parameters": missing_parameters,
        "write_errors": write_errors,
    }


def print_report(elements_count, values, result, collect_mode):
    output.print_md(u"## Заполнение корпуса и секции")
    if collect_mode == "selection":
        output.print_md(u"- Режим: **предварительный выбор**")
    else:
        output.print_md(u"- Режим: **все элементы проекта**")

    output.print_md(u"- Элементов с параметром `{}`: **{}**".format(CODE_PARAMETER, elements_count))
    output.print_md(u"- `{}`: **{}**".format(BUILDING_PARAMETER, values[BUILDING_PARAMETER]))
    output.print_md(u"- `{}`: **{}**".format(SECTION_PARAMETER, values[SECTION_PARAMETER]))
    output.print_md(u"- Записано значений: **{}**".format(result["changed"]))
    output.print_md(u"- Уже было заполнено: **{}**".format(result["already_filled"]))
    output.print_md(u"- Пропущено повторов типоразмера: **{}**".format(result["skipped_type_duplicates"]))
    output.print_md(u"- Не найдено параметров: **{}**".format(len(result["missing_parameters"])))
    output.print_md(u"- Ошибок записи: **{}**".format(len(result["write_errors"])))

    if result["missing_parameters"]:
        output.print_md(u"### Не найдены параметры")
        for element, parameter_name in result["missing_parameters"][:100]:
            output.print_md(u"- {} | `{}`".format(output.linkify(element.Id), parameter_name))
        if len(result["missing_parameters"]) > 100:
            output.print_md(u"- ...и еще **{}**".format(len(result["missing_parameters"]) - 100))

    if result["write_errors"]:
        output.print_md(u"### Ошибки записи")
        for element, parameter_name, reason in result["write_errors"][:100]:
            output.print_md(u"- {} | `{}` | {}".format(output.linkify(element.Id), parameter_name, reason))
        if len(result["write_errors"]) > 100:
            output.print_md(u"- ...и еще **{}**".format(len(result["write_errors"]) - 100))


values_to_write = show_values_form()
if not values_to_write:
    script.exit()

elements_with_code, collect_mode = collect_elements_with_code_parameter()
if not elements_with_code:
    forms.alert(
        u"Не найдено элементов с параметром {}.\n\n"
        u"Если был предварительный выбор, проверь, что выбранные элементы имеют этот параметр у экземпляра или типоразмера.".format(CODE_PARAMETER),
        title=u"Корпус и секция",
    )
    script.exit()

if not confirm_write(len(elements_with_code), values_to_write, collect_mode):
    script.exit()

result = fill_parameters(elements_with_code, values_to_write)
print_report(len(elements_with_code), values_to_write, result, collect_mode)
