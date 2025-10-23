# -*- coding: utf-8 -*-

import clr
from collections import defaultdict

from pyrevit import revit, DB
from pyrevit import script
from pyrevit import forms

doc = __revit__.ActiveUIDocument.Document
output = script.get_output()


def lst_len(lst):
    return sum(1 for _ in lst)

def listdwgs(current_view_only=False):
    dwgs = DB.FilteredElementCollector(doc)\
             .OfClass(DB.ImportInstance)\
             .WhereElementIsNotElementType()\
             .ToElements()

    dwgInst = defaultdict(list)
    if lst_len(dwgs) != 0:
        output.print_md("## СВЯЗАННЫЕ И ИМПОРТИРОВАННЫЕ DWG:")

        for dwg in dwgs:
            if dwg.IsLinked:
                dwgInst["Связанные DWG:"].append(dwg)
            else:
                dwgInst["Импортированные DWG:"].append(dwg)

        for link_mode in dwgInst:
            output.print_md("####{}".format(link_mode))
            for dwg in dwgInst[link_mode]:
                dwg_id = dwg.Id
                dwg_name = \
                    dwg.Parameter[DB.BuiltInParameter.IMPORT_SYMBOL_NAME].AsString()
                dwg_workset = revit.query.get_element_workset(dwg).Name
                dwg_instance_creator = \
                    DB.WorksharingUtils.GetWorksharingTooltipInfo(revit.doc,
                                                                dwg.Id).Creator

                if current_view_only \
                        and revit.active_view.Id != dwg.OwnerViewId:
                    continue

                output.print_md("___")
                output.print_md("- DWG Имя: **{}**\n\n"
                                "- DWG Создан: **{}**\n\n"
                                "- DWG ID: {}\n\n"
                                "- DWG РН: **{}**\n\n"
                                .format(dwg_name,
                                        dwg_instance_creator,
                                        output.linkify(dwg_id),
                                        dwg_workset))
    else:
        output.print_md("## В МОДЕЛЕ ОТСУТСТВУЮТ DWG")  



selected_option = forms.alert("Где искать",
                    options=["Текущая вид",
                            "Во всей модели"])

if selected_option:
    listdwgs(current_view_only=selected_option == "Текущая вид")
