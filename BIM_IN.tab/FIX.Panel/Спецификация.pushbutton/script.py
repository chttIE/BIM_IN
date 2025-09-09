# -*- coding: utf-8 -*-
__title__ = 'Спецификация'
__author__ = 'IliaNistratov'

from pyrevit import forms, script, revit
from Autodesk.Revit.DB import *
from Autodesk.Revit.DB import ElementId, SharedParameterElement

doc = revit.doc
uidoc = revit.uidoc
output = script.get_output()
output.set_width(1200)
output.set_title(__title__)
output.close_others(True)
lfy = output.linkify


# -------------------------------
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# -------------------------------

def get_param_binding_type(doc, definition_obj):
    """Возвращает строку '[Тип]' или '[Экземпляр]' для параметра."""
    binding = doc.ParameterBindings.get_Item(definition_obj)
    if binding:
        if isinstance(binding, InstanceBinding):
            return "[Экземпляр]"
        elif isinstance(binding, TypeBinding):
            return "[Тип]"
    return "[?]"



def get_binding_type(doc, shared_param_elem):
    # Получаем определение
    definition = shared_param_elem.GetDefinition()
    
    # Идём по BindingMap
    it = doc.ParameterBindings.ForwardIterator()
    while it.MoveNext():
        def_in_map = it.Key
        binding = it.Current
        
        # Сравниваем по имени (лучше по GUID)
        if def_in_map.Name == definition.Name:
            if isinstance(binding, InstanceBinding):
                return "По экземпляру"  # Экземпляр
            elif isinstance(binding, TypeBinding):
                return "По типу"      # Типоразмер
    return "Not bound"  # Параметр не привязан

def get_bic_from_schedule(definition):
    cat_id = definition.CategoryId.IntegerValue
    for bic in BuiltInCategory.GetValues(BuiltInCategory):
        if int(bic) == cat_id:
            return bic
    return None

def get_bic_name_from_id(bic_int):
    for bic in BuiltInCategory.GetValues(BuiltInCategory):
        if int(bic) == bic_int:
            return LabelUtils.GetLabelFor(bic)
    return "(Неизвестная категория)"

def get_param_value(param):
    """Возвращает значение параметра с учётом его типа."""
    if not param or not param.HasValue:
        return None
    stype = param.StorageType
    if stype == StorageType.String:
        return param.AsString()
    elif stype == StorageType.Double:
        return param.AsDouble()
    elif stype == StorageType.Integer:
        return param.AsInteger()
    elif stype == StorageType.ElementId:
        return param.AsElementId()
    return None


def set_param_value(param, value):
    """Пытается установить значение параметра, если тип совпадает."""
    if not param or param.IsReadOnly:
        return False
    stype = param.StorageType
    if stype == StorageType.String and isinstance(value, str):
        param.Set(value)
        return True
    elif stype == StorageType.Double and isinstance(value, float):
        param.Set(value)
        return True
    elif stype == StorageType.Integer and isinstance(value, int):
        param.Set(value)
        return True
    elif stype == StorageType.ElementId and isinstance(value, ElementId):
        param.Set(value)
        return True
    return False


def values_equal(val1, val2):
    """Сравнение значений с учётом типов и погрешности для double."""
    if type(val1) != type(val2):
        return False
    if isinstance(val1, float):
        return abs(val1 - val2) < 1e-9
    return val1 == val2


def find_param(el, param_name, doc):
    """
    Ищет параметр сначала у экземпляра, потом у типа.
    Возвращает объект параметра или None.
    """
    param = el.LookupParameter(param_name)
    if not param:
        el_type = doc.GetElement(el.GetTypeId())
        if el_type:
            param = el_type.LookupParameter(param_name)
    return param


def copy_parameter_values(doc, category_bic, old_param_name, new_param_name):
    try:
        elements = list(FilteredElementCollector(doc)
                        .OfCategory(category_bic)
                        .WhereElementIsNotElementType())

        if not elements:
            output.print_md("  Не найдено ни одного элемента в категории для копирования параметров.")
            return

        copied_count = 0
        skipped_count = 0

        with Transaction(doc, "Копирование значений параметров") as t:
            t.Start()

            for el in elements:
                old_param = find_param(el, old_param_name, doc)
                old_value = get_param_value(old_param)
                # print("Значение старого параметр {}".format(old_value))
                if old_value is None:
                    continue

                new_param = find_param(el, new_param_name, doc)
                if not new_param:
                    # print("Параметр новый не найден")
                    continue

                new_value = get_param_value(new_param)
                if values_equal(new_value, old_value):
                    skipped_count += 1
                    continue

                if set_param_value(new_param, old_value):
                    copied_count += 1

            t.Commit()

        output.print_md(
            "-  Скопировано значений: {} | ⏭ Пропущено (уже совпадало): {}".format(copied_count, skipped_count)
        )

    except Exception as e:
        output.print_md("❌ Ошибка при копировании параметров: {}".format(str(e)))


def get_shared_category_parameters(doc, category):
    bindings = doc.ParameterBindings
    iterator = bindings.ForwardIterator()
    iterator.Reset()

    shared_params = []

    while iterator.MoveNext():
        definition = iterator.Key
        binding = iterator.Current

        if not binding.Categories.Contains(category):
            continue

        # Проверяем, существует ли SharedParameterElement с таким именем
        for param_elem in FilteredElementCollector(doc).OfClass(SharedParameterElement):
            if param_elem.Name == definition.Name:
                shared_params.append((definition, param_elem.Id))
                break

    return shared_params

def get_param_binding_type(doc, definition_obj):
    """Возвращает строку '[Тип]' или '[Экземпляр]' для параметра."""
    binding = doc.ParameterBindings.get_Item(definition_obj)
    if binding:
        if isinstance(binding, InstanceBinding):
            return "[Экземпляр]"
        elif isinstance(binding, TypeBinding):
            return "[Тип]"
    return "[?]"


# -------------------------------
# ДОБАВЛЕНИЕ ПАРАМЕТРА В СПЕЦИФИКАЦИЮ
# -------------------------------
def show_and_add_schedule_param(old_name, definition, doc, width_ft=0.5, insert_index=0, exclude_existing=True):
    category_bic = get_bic_from_schedule(definition)

    if category_bic is None:
        output.print_md("  Не удалось определить категорию спецификации.")
        return

    try:
        category = doc.Settings.Categories.get_Item(category_bic)
    except:
        output.print_md("  Не удалось получить категорию из BuiltInCategory.")
        return

    existing_field_ids = list(definition.GetFieldOrder())
    existing_names = set()
    for field_id in existing_field_ids:
        try:
            field = definition.GetField(field_id)
            existing_names.add(field.GetName().strip().lower())
        except:
            pass

    available_params = get_shared_category_parameters(doc, category)

    if exclude_existing:
        filtered_params = [
            (definition_obj, pid) for definition_obj, pid in available_params
            if definition_obj.Name.strip().lower() not in existing_names 
        ]
    else:
        filtered_params = available_params
        
    if not filtered_params:
        output.print_md("  Нет параметров для добавления.")
        return

    # Формируем список с пометками [Тип] или [Экземпляр]
    param_display_names = sorted([
        "{}_{}".format(get_param_binding_type(doc, definition_obj),definition_obj.Name)
        for definition_obj, pid in filtered_params
    ])

    selected_display_name = forms.SelectFromList.show(
        param_display_names,
        multiselect=False,
        title='Заменить {} на:'.format(old_name),
        button_name="Выбрать"
    )

    if not selected_display_name:
        return

    # Убираем пометку из имени при поиске
    selected_clean_name = selected_display_name.split("]_")[1].strip()

    found = next(
        ((d, pid) for d, pid in filtered_params if d.Name.strip() == selected_clean_name),
        None
    )

    if not found:
        output.print_md(" Параметр **{}** не найден среди доступных.".format(selected_clean_name))
        return

    found_definition, found_id = found
    new_name = doc.GetElement(found_id).Name

    try:
        with Transaction(doc, "Добавление поля {}".format(selected_clean_name)) as t:
            t.Start()
            sched_field = definition.InsertField(ScheduleFieldType.Instance, found_id, insert_index)
            sched_field.GridColumnWidth = width_ft
            output.print_md("- Параметр **{}** добавлен в спецификацию.".format(selected_clean_name))
            t.Commit()
    except Exception as e:
        output.print_md("  Ошибка при добавлении параметра: {}".format(str(e)))

    if copy_values:
        copy_parameter_values(doc, category_bic, old_name, new_name)

def get_filtered(definition,field_id):
    filter_count = definition.GetFilterCount()
    filter_infos = []
    for j in range(filter_count):
        sched_filter = definition.GetFilter(j)
        if sched_filter.FieldId == field_id:
            try:
                rule_type = str(sched_filter.FilterType)
                val = sched_filter.GetStringValue() if sched_filter.GetStringValue() else ""
            except:
                val = "(не удалось получить значение)"
            filter_infos.append("Фильтр: тип {} , значение {}".format(rule_type, val))
    if filter_infos: return filter_infos
    else: None

def remove_field(definition,field_id):
    field = definition.GetField(field_id)
    name = field.GetName()
    try:
        with Transaction(doc,"Удаление поля {}".format(name)) as t:
            t.Start()
            definition.RemoveField(field_id)
            output.print_md("- Удаленно поле **{}**".format(name))
            t.Commit()
    except Exception as e: output.print_md("- ОШИБКА при удаленно поле **{}**. Ошибка: {}".format(name,str(e)))


def check_shared_param_in_spf(app, shared_param_element):
    """
    Проверяет, существует ли параметр с таким GUID в текущем ФОП.
    
    :param app: Application (например __revit__.Application)
    :param shared_param_element: SharedParameterElement из проекта
    :return: True/False
    """
    if shared_param_element is None:
        return False

    # GUID из элемента проекта
    param_guid = shared_param_element.GuidValue

    # Подключенный Shared Parameters File
    spf = app.OpenSharedParameterFile()
    if spf is None:
        return False  # ФОП не подключен

    # Обходим все группы и определения
    for group in spf.Groups:
        for definition in group.Definitions:
            try:
                # Каждое определение можно привести к ExternalDefinition
                ext_def = definition
                if hasattr(ext_def, "GUID"):
                    if ext_def.GUID == param_guid:
                        return True
            except:
                continue

    return False

# -------------------------------
# ПАРСИНГ СПЕЦИФИКАЦИИ
# -------------------------------
app = __revit__.Application
def analyze_schedule(view):
    definition = view.Definition
    section = view.GetTableData().GetSectionData(SectionType.Body)
    field_order = definition.GetFieldOrder()
    sort_fields = definition.GetSortGroupFields()

    output.print_md("## Имя: **{}**".format(lfy(view.Id,view.Name)))
    cat_id = definition.CategoryId.IntegerValue
    output.print_md("## Категория: {}".format(get_bic_name_from_id(cat_id)))
    flag = False
    for i, field_id in enumerate(field_order):
        field = definition.GetField(field_id)
        name = field.GetName()
        is_hidden = field.IsHidden
        field_type = str(field.FieldType)
        param_id = field.ParameterId
        insert_index = field.FieldIndex + 1
        width_mm = "-"
        if not is_hidden:
            try:
                width_ft = section.GetColumnWidth(i)
                width_mm = round(width_ft * 304.8, 1)
            except:
                width_mm = "Не смог определить. Принимаю 30"

        param_elem = doc.GetElement(param_id)
        
        if isinstance(param_elem, SharedParameterElement):
            if ignore_adsk and ("ADSK_" in param_elem.Name or check_shared_param_in_spf(app, param_elem)):
                continue
            param_note = ":information: Общий параметр"
            param_display = output.linkify(param_elem.Id, param_elem.Name)
            flag = True 
        else: continue 

        filter_infos = get_filtered(definition,field_id)
        
        sort_info = next(("Сортировка: порядок {}".format('по возрастанию' if sf.SortOrder else 'по убыванию')
                          for sf in sort_fields if sf.FieldId == field_id), "")

        output.print_md("## Необходимо заменить параметр")
        output.print_md("### {}".format(name))
        output.print_md("- **Параметр**: {}".format(param_display))
        output.print_md("- **Тип**: {}".format(param_note))
        output.print_md("- **Ширина**: {} мм".format(width_mm))
        if width_mm == "Не смог определить. Принимаю 30": width_ft = round(30 / 304.8, 1)
        output.print_md("- **Тип поля**: {}".format(get_binding_type(doc, param_elem)))
        output.print_md("- **Скрыт**: {}".format(str(is_hidden)))
        if filter_infos: 
            for info in filter_infos: output.print_md("- **{}**".format(info))
        if sort_info: output.print_md("- **{}**".format(sort_info))
        old_name = name
        
        if add_new_param: show_and_add_schedule_param(old_name,definition, doc, width_ft,insert_index)
        if remove_param: remove_field(definition,field_id)
    if not flag: output.print_md("## Спецификация в порядке")
    output.print_md("---")       
# -------------------------------
# ОСНОВНОЙ ЦИКЛ
# -------------------------------

selected_ids = __revit__.ActiveUIDocument.Selection.GetElementIds() # type: ignore
if selected_ids:
    # views = [doc.GetElement(selected_id) for selected_id in selected_ids]
    views = []
    for selected_id in selected_ids:
        el = doc.GetElement(selected_id)
        if isinstance(el, ViewSchedule):
            views.append(el)
else: 
    views = [doc.ActiveView]
    if not isinstance(views[0], ViewSchedule):
        output.print_md("Активный вид не является спецификацией")
        script.exit()


selected_option, switches = forms.CommandSwitchWindow.show(
    ["Запуск"],
    switches={
        "Добавлять параметры": False,
        "Копировать значения": False,
        "Игнорировать ADSK": True,
        "Удалять замененные": False
    },
    message='Что будет делать?',
    recognize_access_key=False
)

# Если окно закрыли — выход
if selected_option is None:
    script.exit()

# switches уже словарь вида {"Добавлять параметры": True/False, ...}
add_new_param = switches["Добавлять параметры"]
copy_values   = switches["Копировать значения"]
remove_param  = switches["Удалять замененные"]
ignore_adsk   = switches["Игнорировать ADSK"]


for v in views:
    if isinstance(v, ViewSchedule):
        analyze_schedule(v)


