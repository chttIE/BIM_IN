# -*- coding: utf-8 -*-"

from Autodesk.Revit.DB import BuiltInCategory
from pyrevit import script
from sup import collect_elements_on_view,search_and_change_parameters_for_element


doc     = __revit__.ActiveUIDocument.Document   # type: ignore
app     = __revit__.Application                 # type: ignore
doc     = __revit__.ActiveUIDocument.Document   # type: ignore
uidoc   = __revit__.ActiveUIDocument            # type: ignore
output = script.get_output()
lfy = output.linkify
output.close_others(True)

def get_code_model(doc):
    title = doc.Title
    if "K29-К31" in title: return "29-31"
    elif "K28" in title: return "28" 
    elif "K29" in title: return "29"
    elif "K30" in title: return "30"
    elif "K31" in title: return "31"
    else: return "NONE"

name_parameter = "ADSK_Код_Модел"
text = get_code_model(doc)

# Сбор элементов с вида
custom,category = collect_elements_on_view(doc,
    exclude_categories=[BuiltInCategory.OST_SWallRectOpening,
                        BuiltInCategory.OST_Cameras],
    exclude_classes=[],
    preview="off")

output.print_md("___")

bad,g,i = search_and_change_parameters_for_element(doc,custom,name_parameter,text)
if bad:
    output.print_md("ПРОБЛЕМЫ:")
    for i,b in enumerate(bad):
        print("{} Нет параметра у {} - категория {}".format(i+1,lfy(b.Id),b.Category.Name))

output.print_md("Всего элементов {}".format(len(custom)))
if bad:
    output.print_md("У {} Были ошибки ".format(len(bad)))
output.print_md("У {} Было задано ".format(g))
output.print_md("У {} Уже заполнено".format(i))
output.print_md("___")