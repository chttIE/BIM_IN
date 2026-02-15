# -*- coding: utf-8 -*-
__title__ = 'Добавить\nлокальную связь'
__author__ = 'IliaNistratov'
__doc__ =   """Массово добавляет связей с локального диска"""
__highlight__ = 'updated'
from Autodesk.Revit.DB import RevitLinkInstance,FilteredElementCollector as FEC,Transaction
from pyrevit import script, coreutils
from sup import select_file_local
from models import add_link
import datetime
import os
from rpw.ui.forms import (FlexForm, Label, ComboBox, Button, CheckBox)
output = script.get_output()
script.get_output().close_others(all_open_outputs=True)
output.set_title('Добавить связи')

app = __revit__.Application
doc = __revit__.ActiveUIDocument.Document
user = app.Username

links = FEC(doc).OfClass(RevitLinkInstance).ToElements()

# form
def form():
    try:
        Width = 350
        components = [
            Label('Способ размещения связи',
                                            FontSize=14,
                                            Height=30,
                                            Width=Width),
            ComboBox('placement_method',{
                    '2. По началу координат':   1,
                    '3. По базовой точке':      3,
                    '4. По центру координат':   2,
                    '1. По общим координатам':  0},
                                            Width=Width,
                                            FontSize=14,),
            Label("",                       Height=14),
            CheckBox('closed_ws', 'Закрывать все рабочие наборы?',
                                            default=False,
                                            FontSize=14,
                                            Height=30,
                                            Width=Width),
            Button('Добавить связи',
                                            FontSize=14,
                                            Height=30,
                                            Width=Width)]
        
        form = FlexForm("Настройка размещения связей", components)
        form.ShowDialog()
        placement_method = form.values['placement_method']
        closed_ws = form.values["closed_ws"]
        # name_ws = form.values["name_ws"].split(",")
    except:
        script.exit()
    return placement_method,closed_ws

def is_there_link(name_model):
    for link in links:
        if link.Name.split(" ")[0] == name_model:
            return True

#main
placement_method,closed_ws = form()
sel_links = select_file_local()
if sel_links:
    output.print_md("##ДОБАВЛЕНИЕ СВЯЗЕЙ ({})".format(len(sel_links)))
    output.print_md("___")   
    t_timer = coreutils.Timer()  
    output.update_progress(0, len(sel_links))
    with Transaction(doc, 'Добавление связей') as t:
        t.Start() 
        for i, l in enumerate(sel_links):
                name_model = os.path.basename(l)
                if name_model == doc.Title.split("_" + user)[0]:
                    output.print_md("- :cross_mark: Связь **{}** проигнорирована. \
                                    Причина: Попытка загрузить модель в себя же!".format(name_model))
                    continue
                if is_there_link(name_model):
                    output.print_md("- :information: Связь **{}** уже существует".format(name_model))
                    continue
                try:
                    add_link(doc,l,placement_method,closed_ws)
                except Exception as e:
                    output.print_md("- :cross_mark: Ошибка в связи {}. \
                                    Ошибка: {}".format(name_model, str(e)))
                    continue

                output.update_progress(i + 1, len(sel_links))
        t.Commit()

        t_endtime = str(datetime.timedelta(seconds=t_timer.get_time())).split(".")[0]
        output.print_md("___")   
        output.print_md("**Время: {} **".format(t_endtime))
