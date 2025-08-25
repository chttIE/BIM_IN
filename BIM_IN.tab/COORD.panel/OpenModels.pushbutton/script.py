# -*- coding: utf-8 -*-
__title__ = 'Открыть'
__author__ = 'IliaNistratov'
__doc__ =   """
            Массово открывает модели в фоне
            """
__highlight__ = 'updated'
import datetime 
from sup import select_file 
from logIN import lg
from rpw import ui
from models import get_project_path_from_ini,open_model,create_local_model,convert_path
from pyrevit import script,coreutils

output = script.get_output()
script.get_output().close_others(all_open_outputs=True)

doc = __revit__.ActiveUIDocument.Document
app = __revit__.Application
user = app.Username
now = datetime.datetime.now()
current_time = now.strftime("%Y-%m-%d %H:%M")

# form
try:      
	components = [
		ui.forms.flexform.CheckBox('closeallws', 'Закрыть все РН', 
                                    default=True),
        ui.forms.flexform.CheckBox('audit', 'Проверка', 
                                    default=True),
        ui.forms.flexform.CheckBox('not_activate', 'Открыть в фоне', 
                                    default=True),
        ui.forms.flexform.Label("Метод открытия"),
        ui.forms.flexform.ComboBox('detach',    {
                                                '1. Без отсоединения'               :0,
                                                '2. Отсоединить и сохранить РН'     :1,
                                                '3. Отсоединить и не сохранить РН'  :2,
                                                },
                                    FontSize=12),
        ui.forms.Separator(),
        ui.forms.Button('Выбрать')]
	form = ui.forms.FlexForm("Настройка открытия моделей", components)
	form.ShowDialog()
	closeallws = form.values['closeallws']
	audit = form.values["audit"]
	not_activate = form.values["not_activate"]
	detach = form.values["detach"]	
except:
	script.exit()



#main
projectpath = get_project_path_from_ini(doc)
sel_models = select_file()
if sel_models:
    lg(doc,__title__)
    output.print_md("##ОТКРЫТИЕ МОДЕЛЕЙ В ФОНОВОМ РЕЖИМЕ ({})".format(len(sel_models)))
    t_timer = coreutils.Timer()  
    output.update_progress(0, len(sel_models))
    for i, m in enumerate(sel_models):
        output.print_md("___") 
        output.print_md("###Модель: **{}**".format(m))
        targetPath = create_local_model(m, projectpath)

        open_model(targetPath,not_activate,audit,detach,closeallws)
        output.update_progress(i + 1, len(sel_models))
    t_endtime = str(datetime.timedelta(seconds=t_timer.get_time())).split(".")[0]
    output.print_md("___")
    output.print_md("###**Завершено! Время: {} **".format(t_endtime))  