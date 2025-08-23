# -*- coding: utf-8 -*-
#pylint: disable=import-error,invalid-name,broad-except
__title__ = 'Закрыть'
__doc__ =   """
            Закрытие и синхронизация всех открытых моделей кроме текущей. 
            Всегда должна быть одна открытая модель
            """


import datetime
from rpw import ui
from pyrevit import script, forms, coreutils
from logIN import lg
from models import Synchronize_models
doc = __revit__.ActiveUIDocument.Document
output = script.get_output()
script.get_output().close_others(all_open_outputs=True)

# form
try:
	components = [
		ui.forms.flexform.CheckBox('sinhr', 'Синхронизироваться', default=True),
        ui.forms.flexform.CheckBox('flatting', 'Сжать', default=True),
        ui.forms.flexform.CheckBox('reloadlatest', 'Обновить до последней версии', default=False),
        ui.forms.flexform.CheckBox('relinquish_all', 'Отдать все рабочие наборы', default=True),
        ui.forms.flexform.CheckBox('closed', 'Закрыть', default=True),
        ui.forms.flexform.CheckBox('save', 'Сохранить', default=True),
        ui.forms.flexform.CheckBox('comm', 'Оставить комментарий', default=True),
        ui.forms.flexform.Label("Комментарий"),
        ui.forms.flexform.TextBox("comment", Text="Synchronisation from pyIN Panel"),				
        ui.forms.Separator(),
        ui.forms.Button('Выбрать')]
	form = ui.forms.FlexForm("Настройка закрытия моделей", components)
	form.ShowDialog()
	sinhr = form.values['sinhr']
	flatting = form.values["flatting"]
	relinquish_all = form.values["relinquish_all"]
	reloadlatest = form.values["reloadlatest"]
	closed = form.values["closed"]
	save = form.values["save"]
	comm = form.values["comm"]
	comment = form.values["comment"]	
except:
	script.exit()


#main
lg(doc,__title__)
dest_docs = forms.select_open_docs(title='Выбор документов')
if dest_docs:
    output.print_md("##ЗАКРЫТИЕ МОДЕЛЕЙ ({})".format(len(dest_docs)))
    if flatting: output.print_md("- :information: Модель будет **сжата**")
    if comm: output.print_md("- :information: Комментарий к синхронизации: **{}**".format(comment))
    t_timer = coreutils.Timer()  
    output.update_progress(0, len(dest_docs))
    for i, d in enumerate(dest_docs):
        try:
            Synchronize_models(d, sinhr, flatting, relinquish_all,reloadlatest, save, comm,comment)
            if closed:
                d.Close(False)
                output.print_md("-  :white_heavy_check_mark: **Закрытие модели**")
        except Exception as e:
            output.print_md("- :cross_mark: Ошибка в модели {}. Ошибка: {}".format(d.Title, str(e)))
            continue
        output.update_progress(i + 1, len(dest_docs))

    t_endtime = str(datetime.timedelta(seconds=t_timer.get_time())).split(".")[0]
    output.print_md("___")   
    output.print_md("###**Закрытие завершено! Время: {}**".format(t_endtime))
