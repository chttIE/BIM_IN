# -*- coding: utf-8 -*-
__title__ = "Удалить области\nзаливки"
__doc__ = "Удаление типоразмеров и экземпляров области заливки"
__author__ = 'NistratovIlia'

from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from pyrevit import script, forms

doc = __revit__.ActiveUIDocument.Document  # noqa

output = script.get_output()
lfy = output.linkify
output.close_others(all_open_outputs=True)

collector = FilteredElementCollector(doc).OfClass(FilledRegionType).ToElements()

def get_name(el):
    return Element.Name.GetValue(el)

all = []
use = []
not_use = []

# Словарь: типоразмер → список экземпляров
type_to_instances = {}

for c in collector:
    name = get_name(c)
    all.append(name)

    # Получаем зависимые элементы (экземпляры FilledRegion)
    instances = c.GetDependentElements(ElementClassFilter(FilledRegion))
    type_to_instances[c] = instances
    
    if instances:
        output.print_md("use: instance {}".format(lfy(instances)))
        use.append(name)
    else:
        output.print_md("not use: type {}".format(lfy(c.Id)))

        not_use.append(name)

# Показываем диалог выбора
select = forms.SelectFromList.show({
        "Все": all,
        "Используемые": use,
        "Не используемые": not_use,
    },
    multiselect=True,
    group_selector_title='Группы',
    title="Какие удалить?",
    width=500,
    button_name='Удалить'
)

if select:
    with DB.Transaction(doc, "Удаление областей заливки") as t:
        t.Start()
        for region_type in collector:
            name = get_name(region_type)
            if name in select:
                try:
                    output.print_md("Удаляем тип: **{}** ({})".format(name, lfy(region_type.Id)))
                    # Удаляем экземпляры, если есть
                    for inst_id in type_to_instances.get(region_type, []):
                        try:
                            doc.Delete(inst_id)
                            output.print_md("  - Удаляем экземпляр: **{}**".format(lfy(inst_id)))
                        except Exception as e:
                            output.print_md("  - ⚠ Ошибка при удалении экземпляра: {}".format(e))
                    # Удаляем сам тип
                    doc.Delete(region_type.Id)
                except Exception as e:
                    output.print_md("❌ Не удалось удалить тип {}: {}".format(name, e))
        t.Commit()
else:
    script.exit()
