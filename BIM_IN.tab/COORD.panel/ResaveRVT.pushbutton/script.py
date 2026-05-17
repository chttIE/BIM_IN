# -*- coding: utf-8 -*-
__title__ = "ResaveRVT"

import os
import datetime

import clr

clr.AddReference("System")
clr.AddReference("System.Core")
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Controls import DataGridEditingUnit
from System.Net import WebRequest
from System import Uri
from System.Windows.Media import Brushes

from Autodesk.Revit import DB
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
from pyrevit.forms import WPFWindow


# ------------------------------------------------------------
# INIT
# ------------------------------------------------------------

output = script.get_output()
output.close_others(all_open_outputs=True)

app = __revit__.Application
user = app.Username

REST_VERSION = app.VersionNumber

# ------------------------------------------------------------
# HELPERS UI
# ------------------------------------------------------------



# ------------------------------------------------------------
# DATA ITEM
# ------------------------------------------------------------

class ModelItem(object):

    def __init__(self, model_path):
        self.Checked = True
        self.ModelPath = model_path
        self.ModelName = os.path.basename(model_path)

        # Тут храним ТОЛЬКО папку, без имени файла
        self.TargetFolder = ""

        self.IsValid = False
        self.ValidationMessage = "Путь не проверен"


# ------------------------------------------------------------
# REVIT SERVER PATH HELPERS
# ------------------------------------------------------------

def normalize_rs_folder(path):
    if not path:
        return ""

    path = path.strip()
    path = path.replace("\\", "/")

    if path.lower().startswith("rsn://"):
        path = "RSN://" + path[6:]

    while "///" in path:
        path = path.replace("///", "//")

    return path.rstrip("/")


def split_rsn_path(rs_path):
    """
    RSN://server/project/folder
    ->
    server, project/folder
    """

    rs_path = normalize_rs_folder(rs_path)

    if not rs_path.startswith("RSN://"):
        return None, None

    clean = rs_path.replace("RSN://", "", 1)

    parts = clean.split("/")

    if len(parts) < 2:
        return None, None

    server = parts[0]
    folder = "/".join(parts[1:])

    return server, folder


def get_revit_server_hosts():
    try:
        hosts = app.GetRevitServerNetworkHosts()
        return [str(h) for h in hosts]
    except:
        return []


# def validate_rs_folder(rs_folder):
#     """
#     Проверяет, что:
#     1. путь похож на RSN
#     2. сервер доступен в Revit
#     3. папка существует через Revit Server REST API

#     Важно:
#     REST-проверка зависит от версии Revit Server.
#     Например:
#     RevitServerAdminRESTService2021
#     RevitServerAdminRESTService2022
#     """

#     rs_folder = normalize_rs_folder(rs_folder)

#     if not rs_folder:
#         return False, "Путь пустой"

#     if not rs_folder.lower().startswith("rsn://"):
#         return False, "Путь должен начинаться с RSN://"

#     server, folder = split_rsn_path(rs_folder)

#     if not server or not folder:
#         return False, "Неверный формат пути"

#     hosts = get_revit_server_hosts()

#     if hosts:
#         server_exists = False

#         for h in hosts:
#             if h.lower() == server.lower():
#                 server_exists = True
#                 break

#         if not server_exists:
#             return False, "Сервер не найден в Revit Server Network Hosts"

#     try:
#         folder_for_url = folder.replace("/", "|")
#         folder_for_url = Uri.EscapeDataString(folder_for_url)

#         url = (
#             "http://{0}/RevitServerAdminRESTService{1}/"
#             "AdminRESTService.svc/|{2}|/Contents"
#         ).format(server, REST_VERSION, folder_for_url)

#         request = WebRequest.Create(url)
#         request.Method = "GET"
#         request.Timeout = 5000

#         request.Headers.Add("User-Name", user)
#         request.Headers.Add("Revit-User", user)

#         response = request.GetResponse()
#         response.Close()

#         return True, "Папка существует"

#     except Exception as ex:
#         return False, "Папка не найдена или REST недоступен: {}".format(ex)

def validate_rs_folder(rs_folder):
    rs_folder = normalize_rs_folder(rs_folder)

    if not rs_folder:
        return False, "Путь пустой"

    if not rs_folder.lower().startswith("rsn://"):
        return False, "Путь должен начинаться с RSN://"

    server, folder = split_rsn_path(rs_folder)

    if not server or not folder:
        return False, "Неверный формат пути"

    try:
        test_path = rs_folder.rstrip("/") + "/__test__.rvt"
        ModelPathUtils.ConvertUserVisiblePathToModelPath(test_path)
        return True, "Формат пути корректный"

    except Exception as ex:
        return False, "Revit не принимает путь: {}".format(ex)

def build_final_rs_path(item):
    return normalize_rs_folder(item.TargetFolder) + "/" + item.ModelName


# ------------------------------------------------------------
# REVIT OPERATIONS
# ------------------------------------------------------------

def open_local_detached(model_path, audit=False):
    options = OpenOptions()
    options.Audit = audit
    options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets

    ws_config = WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets)
    options.SetOpenWorksetsConfiguration(ws_config)

    return app.OpenDocumentFile(model_path, options)


def save_as_rs_central(doc, rs_path):
    rs_model_path = ModelPathUtils.ConvertUserVisiblePathToModelPath(rs_path)

    save_opts = SaveAsOptions()
    save_opts.OverwriteExistingFile = True
    save_opts.Compact = True

    ws_opts = WorksharingSaveAsOptions()
    ws_opts.SaveAsCentral = True

    save_opts.SetWorksharingOptions(ws_opts)

    doc.SaveAs(rs_model_path, save_opts)

def relinquish_after_save_as_central(doc):
    if not doc.IsWorkshared:
        output.print_md("- :warning: Рабочие наборы не найдены, отдавать нечего")
        return

    try:
        trans_options = DB.TransactWithCentralOptions()

        relinquish_options = DB.RelinquishOptions(True)

        DB.WorksharingUtils.RelinquishOwnership(
            doc,
            relinquish_options,
            trans_options
        )

        output.print_md("- :white_heavy_check_mark: Рабочие наборы отданы")

    except Exception as ex:
        output.print_md(
            "- :warning: Не удалось отдать рабочие наборы: {}".format(ex)
        )

def sync_document(doc):
    if doc.IsWorkshared and not doc.IsFamilyDocument and not doc.IsLinked:
        timer = coreutils.Timer()

        trans_options = DB.TransactWithCentralOptions()
        sync_options = DB.SynchronizeWithCentralOptions()

        relinquish_options = DB.RelinquishOptions(True)
        reload_latest_options = DB.ReloadLatestOptions()
        save_options = DB.SaveOptions()

        sync_options.SetRelinquishOptions(relinquish_options)
        sync_options.Compact = True
        sync_options.Comment = "Synchronisation from BIM_IN"

        doc.Save(save_options)
        doc.ReloadLatest(reload_latest_options)
        doc.Save(save_options)
        doc.SynchronizeWithCentral(trans_options, sync_options)

        endtime = timer.get_time()
        endtime_hms = str(datetime.timedelta(seconds=endtime).seconds)

        output.print_md("{}с. на синхронизацию".format(endtime_hms))

    else:
        output.print_md(
            "Текущий документ не использует совместный режим работы и не может быть синхронизирован"
        )


# ------------------------------------------------------------
# WINDOW
# ------------------------------------------------------------

class ResaveWindow(WPFWindow):

    def __init__(self):
        WPFWindow.__init__(self, "ui.xaml")

        self.items = ObservableCollection[ModelItem]()
        self.grid_models.ItemsSource = self.items

        self.result_items = []

        self.btn_select_models.Click += self.select_models
        self.btn_fill_selected.Click += self.fill_selected
        self.btn_save.Click += self.save_clicked

        self.update_state()

    # --------------------------------------------------------

    def select_models(self, sender, args):
        files = forms.pick_file(
            file_ext="rvt",
            multi_file=True,
            title="Выберите RVT модели"
        )

        if not files:
            return

        if isinstance(files, basestring):
            files = [files]

        self.items.Clear()

        for file_path in files:
            self.items.Add(ModelItem(file_path))

        self.commit_grid_edit()
        self.grid_models.Items.Refresh()
        self.update_state()

    # --------------------------------------------------------

    def fill_selected(self, sender, args):
        folder = forms.ask_for_string(
            default="RSN://server/project/folder",
            prompt="Введите папку Revit Server без имени файла"
        )

        if not folder:
            return

        folder = normalize_rs_folder(folder)

        selected = list(self.grid_models.SelectedItems)

        if not selected:
            selected = list(self.items)

        valid, message = validate_rs_folder(folder)

        for item in selected:
            item.TargetFolder = folder
            item.IsValid = valid
            item.ValidationMessage = message

        self.commit_grid_edit()
        self.grid_models.Items.Refresh()
        self.repaint_textboxes()
        self.update_state()

    # --------------------------------------------------------

    def checkbox_clicked(self, sender, args):
        self.commit_grid_edit()

        row = sender.DataContext
        new_value = bool(sender.IsChecked)

        selected = list(self.grid_models.SelectedItems)

        if row in selected and len(selected) > 1:
            for item in selected:
                item.Checked = new_value
        else:
            row.Checked = new_value

        self.update_state()
    # --------------------------------------------------------

    def target_folder_lost_focus(self, sender, args):
        item = sender.DataContext

        folder = normalize_rs_folder(sender.Text)

        item.TargetFolder = folder

        valid, message = validate_rs_folder(folder)

        item.IsValid = valid
        item.ValidationMessage = message

        if valid:
            sender.BorderBrush = Brushes.LightGray
            sender.Background = Brushes.White
            sender.ToolTip = message
        else:
            sender.BorderBrush = Brushes.Red
            sender.Background = Brushes.MistyRose
            sender.ToolTip = message

        self.update_state()

    # --------------------------------------------------------

    def repaint_textboxes(self):
        """
        DataGrid виртуализирует строки, поэтому полного доступа ко всем TextBox нет.
        Refresh достаточно для обновления данных.
        Цвет у ручного ввода обновляется после LostFocus.
        """
        pass

    # --------------------------------------------------------

    def update_state(self):
        checked_count = 0
        invalid_count = 0

        for item in self.items:
            if item.Checked:
                checked_count += 1

                if not item.IsValid:
                    invalid_count += 1

        self.btn_save.IsEnabled = checked_count > 0 and invalid_count == 0

        self.txt_status.Text = (
            "Моделей выбрано: {} | Ошибок пути: {}"
            .format(checked_count, invalid_count)
        )

    # --------------------------------------------------------

    def save_clicked(self, sender, args):
        result = []

        for item in self.items:
            if item.Checked:
                valid, message = validate_rs_folder(item.TargetFolder)

                item.IsValid = valid
                item.ValidationMessage = message

                if valid:
                    result.append(item)

        self.commit_grid_edit()
        self.grid_models.Items.Refresh()
        self.update_state()

        if not result:
            forms.alert("Нет моделей с валидными путями.")
            return

        for item in self.items:
            if item.Checked and not item.IsValid:
                forms.alert(
                    "Есть модели с неверным путем.\n\n"
                    "Проверьте красные строки."
                )
                return

        self.result_items = result
        self.Close()

    def commit_grid_edit(self):
        try:
            self.grid_models.CommitEdit(DataGridEditingUnit.Cell, True)
            self.grid_models.CommitEdit(DataGridEditingUnit.Row, True)
        except:
            pass
# ------------------------------------------------------------
# UI START
# ------------------------------------------------------------

window = ResaveWindow()
window.ShowDialog()

selected_items = window.result_items

if not selected_items:
    script.exit()


# ------------------------------------------------------------
# PROCESS
# ------------------------------------------------------------

output.print_md(
    "## Экспорт моделей в Revit Server ({})".format(len(selected_items))
)

output.update_progress(0, len(selected_items))
timer = coreutils.Timer()

for i, item in enumerate(selected_items):
    output.update_progress(i + 1, len(selected_items))

    output.print_md("___")
    output.print_md("### Модель: **{}**".format(item.ModelPath))

    doc = None

    try:
        model_mp = ModelPathUtils.ConvertUserVisiblePathToModelPath(item.ModelPath)

        output.print_md("- :white_heavy_check_mark: Открытие модели")

        doc = open_local_detached(model_mp, audit=False)

        output.print_md("- :white_heavy_check_mark: Открыта с отсоединением от central")

        name_model = doc.Title.split("_отсоединено")[0]

        final_rs_path = normalize_rs_folder(item.TargetFolder) + "/" + name_model + ".rvt"

        output.print_md(
            "- :white_heavy_check_mark: Папка назначения: {}".format(item.TargetFolder)
        )

        output.print_md(
            "- :white_heavy_check_mark: Полный путь сохранения: {}".format(final_rs_path)
        )

        save_as_rs_central(doc, final_rs_path)

        output.print_md("- :white_heavy_check_mark: Сохранено как ФХ на Revit Server")

        relinquish_after_save_as_central(doc)

        doc.Close(False)
        doc = None

        output.print_md("- :white_heavy_check_mark: Модель закрыта")

    except Exception as ex:
        output.print_md("- :cross_mark: Ошибка: {}".format(ex))

        try:
            if doc:
                doc.Close(False)
        except:
            pass


# ------------------------------------------------------------
# DONE
# ------------------------------------------------------------

total_time = str(datetime.timedelta(seconds=timer.get_time())).split(".")[0]

output.print_md("___")
output.print_md("### Готово. Время выполнения: **{}**".format(total_time))