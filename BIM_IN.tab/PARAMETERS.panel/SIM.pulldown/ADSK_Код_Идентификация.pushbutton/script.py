# -*- coding: utf-8 -*-"
from pyrevit import script
from Autodesk.Revit.DB import BuiltInCategory, BuiltInParameter, Transaction
from sup import collect_elements_on_view


doc     = __revit__.ActiveUIDocument.Document   # type: ignore
app     = __revit__.Application                 # type: ignore
doc     = __revit__.ActiveUIDocument.Document   # type: ignore
uidoc   = __revit__.ActiveUIDocument            # type: ignore
output = script.get_output()
lfy = output.linkify
output.close_others(True)
name_parameter = "ADSK_Код_Идентификация"
name_parameter1 = "Код_семейства_Revit"
name_parameter2 = "ADSK_Код_Дисциплин"
name_parameter3 = "ADSK_Код_Модел"
name_parameter4 = "ADSK_Код_Расположения"
separator = "_"




def changes(elements):
    bad= []
    g,i = 0,0
    with Transaction(doc,"Заполнение параметра") as t:

        t.Start()
        try:
            c=0
            for el in elements: 
                p_in = el.LookupParameter(name_parameter)
                p_1 = el.LookupParameter(name_parameter1)
                p_2 = el.LookupParameter(name_parameter2)
                p_3 = el.LookupParameter(name_parameter3)
                p_4 = el.LookupParameter(name_parameter4)
                if not p_in:
                    print("ПАРАМЕТР ДЛЯ ЗАПОЛНЕНИЯ ОТСУТСТВУЕТ {} ОТРЕДАКТИРУЙТЕ ПАРАМЕТРЫ!".format(lfy(el.Id)))
                    continue
                if not (p_1 and p_2 and p_3 and p_4):
                    print("ОДИН ИЗ ПАРАМЕТРОВ {} {} {} {} ОТСУТСТВУЕТ У ЭЛЕМЕНТА {} ОТРЕДАКТИРУЙТЕ ПАРАМЕТРЫ!".format(name_parameter1,name_parameter2,name_parameter3,name_parameter4,lfy(el.Id)))
                    continue
                text = "{}+{}+{}+{}".format(p_1.AsString(),p_2.AsString(),p_3.AsString(),p_4.AsString())
                if not p_in:
                    print("ПАРАМЕТР ДЛЯ ЗАПОЛНЕНИЯ ОТСУТСТВУЕТ {} ОТРЕДАКТИРУЙТЕ ПАРАМЕТРЫ!".format(lfy(el.Id)))
                    continue
                if p_in.IsReadOnly: 
                    print("ПАРАМЕТР ДЛЯ ЗАПОЛНЕНИЯ ДОСТУПЕН ТОЛЬКО ДЛЯ ЧТЕНИЯ {} ОТРЕДАКТИРУЙТЕ ПАРАМЕТРЫ!".format(lfy(el.Id)))
                    continue
                if p_in.AsString() == text:
                    i+=1
                    continue
                else:
                    p_in.Set(text)
                    g+=1
                    c+=1

            # doc.Regenerate()

        except Exception as e:
            print("ОШИБКУ У {} КАТЕГОРИИ {} ПРОВЕРЬТЕ ЭЛЕМЕНТ И ПАРАМЕТР!".format(lfy(el.Id),el.Category.Name))


        t.Commit()
    return bad,g,i





# Сбор элементов с вида
custom,category = collect_elements_on_view(doc,
    exclude_categories=[BuiltInCategory.OST_SWallRectOpening,
                        BuiltInCategory.OST_Cameras],
    exclude_classes=[],
    preview="off")
output.freeze()
output.print_md("___")

# for el in custom:
#     level = get_level(el)


bad,g,i = changes(custom)
# if bad:
#     output.print_md("ПРОБЛЕМЫ:")
#     for i,b in enumerate(bad):
#         print(bad)

output.print_md("Всего элементов {}".format(len(custom)))
if bad:
    output.print_md("У {} Были ошибки ".format(len(bad)))
output.print_md("У {} Было задано ".format(g))
output.print_md("У {} Уже заполнено".format(i))
output.print_md("___")
output.unfreeze()