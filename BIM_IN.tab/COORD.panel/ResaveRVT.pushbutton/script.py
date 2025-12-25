# -*- coding: utf-8 -*-
__title__ = "ResaveRVT"

import os
import datetime

from Autodesk.Revit.DB import (
    OpenOptions,
    DetachFromCentralOption,
    WorksetConfiguration,
    WorksetConfigurationOption,
    SaveAsOptions,
    WorksharingSaveAsOptions,
    ModelPathUtils
)

from pyrevit import forms, script, coreutils

# ------------------------------------------------------------
# INIT
# ------------------------------------------------------------
output = script.get_output()
output.close_others(all_open_outputs=True)

app = __revit__.Application
user = app.Username

# ------------------------------------------------------------
# OPEN LOCAL MODEL WITH DETACH
# ------------------------------------------------------------
def open_local_detached(model_path, audit=False):
    options = OpenOptions()
    options.Audit = audit
    options.DetachFromCentralOption = (DetachFromCentralOption.DetachAndPreserveWorksets)
    ws_config = WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets)
    options.SetOpenWorksetsConfiguration(ws_config)
    return app.OpenDocumentFile(model_path, options)

# ------------------------------------------------------------
# SAVE AS CENTRAL TO REVIT SERVER
# ------------------------------------------------------------
def save_as_rs_central(doc, rs_path):
    rs_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(rs_path)
    save_opts = SaveAsOptions()
    save_opts.OverwriteExistingFile = True
    save_opts.Compact = True
    ws_opts = WorksharingSaveAsOptions()
    ws_opts.SaveAsCentral = True
    save_opts.SetWorksharingOptions(ws_opts)
    doc.SaveAs(rs_model_path, save_opts)

# ------------------------------------------------------------
# EXTRACT FOLDER IN NAME MODEL 
# ------------------------------------------------------------

def get_rs_folder_from_filename(filename):
    """
    11692_HVAC_VN_GENPRO_R22_B.rvt → HVAC
    """
    name = os.path.splitext(os.path.basename(filename))[0]
    parts = name.split("_")

    if len(parts) < 2:
        return "UNSORTED"

    return parts[1]

# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------
sel_folder = forms.pick_folder(title="Выбор папки с локальными RVT")
if not sel_folder:
    script.exit()

# берем только RVT
all_models = [os.path.join(sel_folder, f)
    for f in os.listdir(sel_folder)
    if f.lower().endswith(".rvt")]

if not all_models:
    forms.alert("В папке нет RVT файлов")
    script.exit()

# выбор моделей
display = {os.path.basename(p): p for p in all_models}

selected = forms.SelectFromList.show(
    sorted(display.keys()),
    title="Выберите локальные копии для переноса",
    multiselect=True,
    button_name="Перенести в Revit Server"
)

if not selected:
    script.exit()

sel_models = [display[name] for name in selected]

output.print_md(
    "## Экспорт моделей в Revit Server ({})".format(len(sel_models))
)

output.update_progress(0, len(sel_models))
timer = coreutils.Timer()

# ------------------------------------------------------------
# PROCESS
# ------------------------------------------------------------
for i, model in enumerate(sel_models):
    output.update_progress(i + 1, len(sel_models))
    output.print_md("___")
    output.print_md("### Модель: **{}**".format(model))

    try:
        model_mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(model)
        output.print_md(
            "- :white_heavy_check_mark: Открытие модели"
        )
        doc = open_local_detached(model_mp, audit=False)
        name_model = doc.Title.split("_отсоединено")[0]
        output.print_md(
            "- :white_heavy_check_mark: Открыта с отсоединением от central"
        )

        # !!! ВАЖНО: путь Revit Server должен существовать
        rs_subfolder = get_rs_folder_from_filename(model)
        output.print_md(
            "- :white_heavy_check_mark: Выбрана папка {}".format(rs_subfolder)
        )
        rs_path = (
            "RSN://hueugfeno5lf.rsnbim.ru/"
            "Пыхтино/01. В работе/{}/{}.rvt"
            .format(rs_subfolder, name_model)
        )
        output.print_md(
            "- :white_heavy_check_mark: Полный путь {}".format(rs_path)
        )
        save_as_rs_central(doc, rs_path)

        output.print_md(
            "- :white_heavy_check_mark: Сохранено как ФХ на Revit Server"
        )

        doc.Close(False)
        output.print_md(
            "- :white_heavy_check_mark: Модель закрыта"
        )
    except Exception as ex:
        output.print_md(
            "- :cross_mark: Ошибка: {}".format(ex)
        )

# ------------------------------------------------------------
# DONE
# ------------------------------------------------------------
total_time = str(
    datetime.timedelta(seconds=timer.get_time())
).split(".")[0]

output.print_md("___")
output.print_md("### Готово. Время выполнения: **{}**".format(total_time))
