# -*- coding: utf-8 -*-"

from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import script
from sup import collect_elements_on_view,search_and_change_parameters_for_element
import csv

doc     = __revit__.ActiveUIDocument.Document   # type: ignore
app     = __revit__.Application                 # type: ignore
doc     = __revit__.ActiveUIDocument.Document   # type: ignore
uidoc   = __revit__.ActiveUIDocument            # type: ignore
output = script.get_output()
lfy = output.linkify
output.close_others(True)


def get_text(doc):
    title = doc.Title
    if "AR_FC" in title: return "AR_FC"
    elif "AR_AR" in title: return "AR_AR" 
    elif "AR_AI" in title: return "AR_AI"
    else: return "NONE"

name_parameter = "Код_семейства_Revit"
text = get_text(doc)

# Сбор элементов с вида
custom,category = collect_elements_on_view(doc,
    exclude_categories=[BuiltInCategory.OST_SWallRectOpening],
    exclude_classes=[],
    preview="off")

output.print_md("___")

bad,g,i = search_and_change_parameters_for_element(doc,custom,name_parameter,text)
if bad:
    output.print_md("ПРОБЛЕМЫ:")
    for b in bad:
        print("Нет параметра у {} - категория {}".format(lfy(b.Id),b.Category.Name))

output.print_md("Всего элементов {}".format(len(custom)))
output.print_md("У {} Было задано ".format(g))
output.print_md("У {} Уже заполнено".format(i))
output.print_md("___")