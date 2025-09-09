# -*- coding: utf-8 -*-


from Autodesk.Revit.DB import FilteredWorksetCollector as FWC, Transaction, WorksetId, WorksetKind
from pyrevit import script
from rpw import ui


output = script.get_output()
script.get_output().close_others(all_open_outputs=True)
lfy = output.linkify
doc = __revit__.ActiveUIDocument.Document  # type: ignore
try:
    components = [
        ui.forms.flexform.Label("Что искать"),
        ui.forms.flexform.TextBox("comment1", Text="RD"),
        ui.forms.flexform.Label("На что заменить"),
        ui.forms.flexform.TextBox("comment2", Text="PD"),
        ui.forms.Separator(),
        ui.forms.Button('Выбрать')
    ]
    form = ui.forms.FlexForm("Заменить имя РН", components)
    form.ShowDialog()

    name_old = str(form.values["comment1"])
    name_new = str(form.values["comment2"])
except:
    script.exit()


worksets = FWC(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
table = doc.GetWorksetTable()
count=0
with Transaction(doc, 'Переименование РН') as t:
    t.Start()
    for w in worksets:
        name_ws = w.Name
        if name_old in name_ws:
            new_name = name_ws.replace(name_old,name_new)
            id_of_workset = WorksetId(w.Id.IntegerValue)
            table.RenameWorkset(doc, id_of_workset, new_name)
            output.print_md("- Переименовал: {} -> {}".format(name_ws,new_name))
            count+=1
    t.Commit()
    
