# -*- coding: utf-8 -*-

from Autodesk.Revit import DB
from pyrevit import forms, revit, script

output = script.get_output()
output.close_others(True)



def close_selected_worksets(d, prefixes=None):

    link_instances = (
        DB.FilteredElementCollector(d)
        .OfClass(DB.RevitLinkInstance)
        .ToElements()
    )

    if not link_instances:
        output.print_md("Связи не найдены")
        return

    for inst in link_instances:
        link_type = d.GetElement(inst.GetTypeId())
        ext_ref = link_type.GetExternalFileReference()
        l_d = inst.GetLinkDocument()
        if not l_d:
            output.print_md("- Связь {} выгружена и будет пропущена".format(inst.Name))
            continue
        # Получаем список рабочих наборов связи
        output.print_md("___")
        output.print_md("- Связь {} ".format(inst.Name))
        worksets = DB.FilteredWorksetCollector(l_d)\
            .OfKind(DB.WorksetKind.UserWorkset)\
            .ToWorksets()
        for ws in worksets:
            output.print_md(">> {}".format(ws.Name))
        
        ws_config = DB.WorksetConfiguration(
            DB.WorksetConfigurationOption.OpenAllWorksets
        )

        closed = []
        wss_for_open = []
        for ws in worksets:
            output.print_md("- {}".format(ws.Name))
            # # если имя содержит один из префиксов — НЕ открываем
            # if any(p.lower() in ws.Name.lower() for p in prefixes):
            #     closed.append(ws.Name)
            #     continue
            wss_for_open.append(ws)
            
        ws_config.Open([ws.Id for ws in wss_for_open]) 
        link_type.LoadFrom(ext_ref.GetAbsolutePath(), ws_config)


        if closed:
            output.print_md(
                u"- {} → закрыто: {}".format(
                    inst.Name,
                    ", ".join(closed)
                )
            )


# ------------------------------------------------------------
# ВВОД
# ------------------------------------------------------------

# name_for_closed = forms.ask_for_string(
#     default="(50), #",
#     title="Префиксы РН для закрытия",
#     prompt="Введите через запятую"
# )

# if not name_for_closed:
#     script.exit()

# prefixes = [x.strip() for x in name_for_closed.split(",") if x.strip()]

dest_docs = forms.select_open_docs(title='Выбор документов',check_more_than_one=False)

if not dest_docs:
    script.exit()

output.print_md("## Закрытие рабочих наборов ({})".format(len(dest_docs)))

for d in dest_docs:
    output.print_md("### {}".format(d.Title))
    close_selected_worksets(d, prefixes)