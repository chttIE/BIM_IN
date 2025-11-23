# -*- coding: utf-8 -*-
__title__ = 'Открыть'
__author__ = 'IliaNistratov'
__doc__ =   """Массово открывает модели в фоне"""
__highlight__ = 'updated'


import datetime
import os 
from Autodesk.Revit.DB import Document, ModelPathUtils
from rpw import ui
from pyrevit import forms, script,coreutils
from models import (get_project_path_from_ini,
                    open_model,
                    create_local_model, save_as_model,
                    select_file_local)

output = script.get_output()
script.get_output().close_others(all_open_outputs=True)

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
user = app.Username
now = datetime.datetime.now()
current_time = now.strftime("%Y-%m-%d %H:%M")



folder_path = forms.pick_folder(title="Выбор папки для сохранения RVT", owner=None)

#main
projectpath = get_project_path_from_ini(doc)
sel_models = select_file_local()
if sel_models:
    output.print_md("##ЭКCПОРТ RVT МОДЕЛЕЙ В ФОНОВОМ РЕЖИМЕ ({})".format(len(sel_models)))
    t_timer = coreutils.Timer()  
    output.update_progress(0, len(sel_models))
    for i, m in enumerate(sel_models):
        output.print_md("___") 
        output.print_md("###Модель: **{}**".format(m))
        targetPath = create_local_model(m, projectpath)
        d = open_model(targetPath,0,False,3,True)
        d_name = str(d.Title).replace("_отсоединено","")
        d_name = d_name.replace("_{}".format(str(user)),"")
        d_name = "{}.rvt".format(d_name)
        print(d_name)
        model_path = os.path.join(folder_path,d_name)
        print(model_path)
        targetPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(model_path)
        save_as_model(d,
                  targetPath,
                  compact=True,
                  maxbackups=1,
                  overwrite=True)
        d.Close(False)
        output.update_progress(i + 1, len(sel_models))
    t_endtime = str(datetime.timedelta(seconds=t_timer.get_time())).split(".")[0]
    output.print_md("___")
    output.print_md("###**Завершено! Время: {} **".format(t_endtime))  