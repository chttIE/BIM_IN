# -*- coding: utf-8 -*-
__title__ = "EXPORT"
__author__ = 'NistratovIlia'
__doc__ = """Массовый экспорт моделей в формат NWC"""
context = "zero-doc"

import os
import datetime
import codecs
import re
import System.Windows # type: ignore

import clr
clr.AddReference("PresentationFramework")
from System.Windows import Window # type: ignore
from pyrevit import script,coreutils,forms
from Autodesk.Revit.DB import   DetachFromCentralOption, ModelPathUtils, NavisworksCoordinates, NavisworksExportScope, NavisworksParameters, \
                                OpenOptions, WorksetConfiguration, \
                                WorksetConfigurationOption, WorksharingUtils,\
                                NavisworksExportOptions, FilteredElementCollector as FEC, \
                                OptionalFunctionalityUtils, View3D

from rpw.ui.forms import (FlexForm, Label, ComboBox, TextBox, Separator, Button,CheckBox)

output = script.get_output()
script.get_output().close_others(all_open_outputs=True)

doc = __revit__.ActiveUIDocument.Document # type: ignore
app = __revit__.Application # type: ignore
user = app.Username
now = datetime.datetime.now()
current_time = now.strftime("%Y-%m-%d %H:%M")
# output.set_width(2000)


#Проверка наличия утилиты для экспорта
if not OptionalFunctionalityUtils.IsNavisworksExporterAvailable(): 
    forms.alert("Отсутствует модуль экпорта NWC. Обратитесь к координатору",exitscript=True)


# name_view = "Navisworks"
prefix = ""
suffix = ""
folder_path = None
files_path = None
files_txt = None



class ButtonClass(Window):
    @staticmethod
    def SelectFolder(sender, e):
        global folder_path
        folder_path = forms.pick_folder(title="Выбор папки для сохранения NWC", owner=None)
        if folder_path: output.print_md("Путь куда сохранять файлы {}".format(folder_path))
        return(folder_path)

    @staticmethod
    def ShowNwcOptions(sender=None, args=None):
        ok = form_nwc_options()
        # тут можно вывести тост/лог при успехе
        output.print_md(":gear: Настройки NWC обновлены") if ok else None
    @staticmethod
    def SelectFiles(sender, e):
        global files_path
        files_path = forms.pick_file(title="Выбор RVT моделей", file_ext='rvt', multi_file=True)
        if files_path: 
            output.print_md("Выбранные модели:".format(files_path))
            for m in files_path:
                output.print_md("- {}".format(m))
        return files_path

    @staticmethod
    def SelectTxt(sender, e):
        global files_txt
        files_txt = forms.pick_file(title="Выбор txt файла с адресами", file_ext='txt', multi_file=False)
        if files_txt:
            files_txt = select_file_in_txt(files_txt)
            if files_txt:
                output.print_md("Выбранные модели:")
                for m in files_txt:
                    output.print_md("- {}".format(m))
        return files_txt

# form
# ---- Глобальные/модульные настройки NWC с дефолтами ----
NWC_OPTIONS = {
    'ConvertElementProperties': True,
    'ExportRoomGeometry': True,
    'ExportRoomAsAttribute': True,
    'ExportUrls': True,
    'ExportLinks': False,
    'ConvertLinkedCADFormats': True,
    'ExportElementIds': True,
    'ExportParts': False,
    'FindMissingMaterials': True,
    'ConvertLights': False,
    'DivideFileIntoLevels': True,
    'Coordinates': NavisworksCoordinates.Shared,
    'ExportScope': NavisworksExportScope.View,
    'FacetingFactor': 25,
    'Parameters': NavisworksParameters.All,
}

# ---- Форма настроек NWC (вызывается по кнопке) ----
def form_nwc_options():
    Width = 320
    FontSize = 14
    FontFamily = System.Windows.Media.FontFamily('Consolas')

    coords_items = [
        "Общие координаты (Shared)",
        "Внутренние координаты (Internal)",
        "Координаты проекта (Project)"
    ]
    scope_items = [
        "Текущий вид",
        "Вся модель",
        "Выбранные элементы"
    ]

    components = [
        Label("Координаты:", Height=26),
        ComboBox('coords', coords_items, default=coords_items[{NavisworksCoordinates.Shared:0, 
                                                               NavisworksCoordinates.Internal:1}[NWC_OPTIONS['Coordinates']]]),
        Label("Область экспорта:", Height=26),
        ComboBox('scope', scope_items, default=scope_items[{NavisworksExportScope.View:0, 
                                                            NavisworksExportScope.Model:1}[NWC_OPTIONS['ExportScope']]]),
        Label("Faceting Factor:", Height=26),
        TextBox('faceting', Text=str(NWC_OPTIONS['FacetingFactor']), Height=28),

        Separator(),

        CheckBox('prop',        'Конвертировать свойства элементов',       default=NWC_OPTIONS['ConvertElementProperties']),
        CheckBox('room_geom',   'Экспортировать геометрию помещений',      default=NWC_OPTIONS['ExportRoomGeometry']),
        CheckBox('room_attr',   'Экспортировать помещения как атрибуты',   default=NWC_OPTIONS['ExportRoomAsAttribute']),
        CheckBox('urls',        'Экспортировать URL (ссылки)',             default=NWC_OPTIONS['ExportUrls']),
        CheckBox('links',       'Экспортировать Revit-ссылки',             default=NWC_OPTIONS['ExportLinks']),
        CheckBox('linked_cad',  'Конвертировать связанные CAD',            default=NWC_OPTIONS['ConvertLinkedCADFormats']),
        CheckBox('elem_ids',    'Экспортировать ElementId',                default=NWC_OPTIONS['ExportElementIds']),
        CheckBox('parts',       'Экспортировать детали (Parts)',           default=NWC_OPTIONS['ExportParts']),
        CheckBox('miss_mat',    'Искать отсутствующие материалы',          default=NWC_OPTIONS['FindMissingMaterials']),
        CheckBox('lights',      'Конвертировать источники света',          default=NWC_OPTIONS['ConvertLights']),
        CheckBox('levels',      'Делить файл по уровням',                  default=NWC_OPTIONS['DivideFileIntoLevels']),

        Separator(),
        Button('OK', Height=30),
    ]

    ff = FlexForm("Настройки Navisworks Export", components)
    ff.ShowDialog()
    v = ff.values
    if not v:  # окно закрыли крестиком
        return False

    # Парсинг
    try:
        faceting = max(1, int(v.get('faceting', NWC_OPTIONS['FacetingFactor'])))
    except:
        faceting = NWC_OPTIONS['FacetingFactor']

    coords_map = {
        "Общие координаты (Shared)": NavisworksCoordinates.Shared,
        "Внутренние координаты (Internal)": NavisworksCoordinates.Internal,
    }
    scope_map = {
        "Текущий вид": NavisworksExportScope.View,
        "Вся модель":  NavisworksExportScope.Model,
    }

    # Обновляем глобальный словарь
    NWC_OPTIONS.update({
        'ConvertElementProperties': v.get('prop', True),
        'ExportRoomGeometry':       v.get('room_geom', True),
        'ExportRoomAsAttribute':    v.get('room_attr', True),
        'ExportUrls':               v.get('urls', True),
        'ExportLinks':              v.get('links', False),
        'ConvertLinkedCADFormats':  v.get('linked_cad', True),
        'ExportElementIds':         v.get('elem_ids', True),
        'ExportParts':              v.get('parts', False),
        'FindMissingMaterials':     v.get('miss_mat', True),
        'ConvertLights':            v.get('lights', False),
        'DivideFileIntoLevels':     v.get('levels', True),
        'Coordinates':              coords_map.get(v['coords'], NWC_OPTIONS['Coordinates']),
        'ExportScope':              scope_map.get(v['scope'],  NWC_OPTIONS['ExportScope']),
        'FacetingFactor':           faceting,
        'Parameters':               NavisworksParameters.All,
    })
    return True



# ---- Главная форма: добавили кнопку "Настройки экспорта NWC" ----

# ---- Главная форма: добавили кнопку "Настройки экспорта NWC" ----
def form_main():
    Width = 300
    FontSize = 14
    FontFamily = System.Windows.Media.FontFamily('Consolas')

    components = [
        Button('Выбрать модели RVT из папки', on_click=ButtonClass.SelectFiles),
        Label('                                   Или', Height=28),
        Button('Выбрать txt c адресами', on_click=ButtonClass.SelectTxt),

        Separator(),
        Label('Имя вида для экспорта:', Height=28),
        TextBox("name_view", Text="Navisworks", Height=28),
        Separator(),

        CheckBox('open_wss', 'Открыть рабочие наборы', default=True, Height=28),
        Label("Слова в именах рабочих наборов (через запятую)", Height=30),
        TextBox('name_ws', Text="(20)_,(30)_,(40)_", TextWrapping=System.Windows.TextWrapping.Wrap,
                AcceptsTab=True, AcceptsReturn=True, Multiline=True, Height=50),

        Label("Префикс к имени", Height=30),
        TextBox('prefix', Text="", TextWrapping=System.Windows.TextWrapping.Wrap, Height=28),

        Label("Суффикс к имени", Height=30),
        TextBox('suffix', Text="", TextWrapping=System.Windows.TextWrapping.Wrap, Height=28),

        Separator(),

        # <-- вот наша кнопка настроек; 
        Button('Настройки экспорта NWC…', on_click=ButtonClass.ShowNwcOptions),
        CheckBox('save_html', 'Сохранить отчет', default=False, Height=28),
        CheckBox('open_folder', 'Открыть папку после экспорта', default=True, Height=28),

        Separator(),
        Button('Куда сохранить NWC', on_click=ButtonClass.SelectFolder),

        Button('EXPORT', Height=30),
    ]

    ff = FlexForm("Настройка экспорта", components)
    ff.ShowDialog()
    v = ff.values or {}
    # разбор значений основной формы
    name_view   = v.get("name_view", "Navisworks")
    open_wss    = v.get("open_wss", True)
    name_ws     = v.get("name_ws", "").split(",") if v.get("name_ws") else []
    save_html   = v.get("save_html", False)
    open_folder = v.get("open_folder", True)
    prefix      = v.get("prefix", "")
    suffix      = v.get("suffix", "")
    return name_view, open_wss, name_ws, save_html, open_folder, prefix, suffix



def get_ViewExportId(d,name_view):
    for view in FEC(d).OfClass(View3D):
        if view.Name == name_view:
            output.print_md("-  :white_heavy_check_mark: Вид для экспорта найден: **{}**".format(name_view))
            return view.Id
    else:  
        output.print_md("- :cross_mark: Вид для экспорта отсутствует: **{}**".format(name_view))
        return False
    
def export_NWC(d, view_id, folder_path, prefix="", suffix=""):
    # имя файла
    name_model = d.Title.split("_отсоединено")[0]
    name_model = "{1}{0}{2}".format(name_model, prefix, suffix)

    # собрать опции из NWC_OPTIONS
    nweo = NavisworksExportOptions()
    nweo.ViewId = view_id
    for k, v in NWC_OPTIONS.items():
        setattr(nweo, k, v)

    e_timer = coreutils.Timer() 
    d.Export(folder_path,name_model,nweo)
    e_endtime = str(datetime.timedelta(seconds=e_timer.get_time())).split(".")[0]
    output.print_md("-  :white_heavy_check_mark: **Экспорт модели завершен! Время: {}**".format(e_endtime))

def get_ws_for_open(mp, name_ws):
    ws_for_open = []
    wss_link = WorksharingUtils.GetUserWorksetInfo(mp)
    for ws in wss_link:
        for name in name_ws:
            if "{}".format(name) in ws.Name and "Отвер" not in ws.Name:
                ws_for_open.append(ws)
    return ws_for_open


def open_model( path,
                activate = True,
                audit = True,
                detach = 1,
                closeallws = True,
                log=2):
    # Задаем настройки открытия моделей
    options = OpenOptions()
    
    if closeallws == True: 
        workset_config = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        if log > 0: output.print_md("- :information: Рабочие наборы будут **закрыты**")
    elif closeallws == False: 
        workset_config = WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets)
        if log > 0: output.print_md("- :information: Рабочие наборы будут **открыты**")
    if isinstance(closeallws,list):
        workset_config = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        wss_for_open = get_ws_for_open(path,closeallws)
        wss_name_for_open = [ws.Name for ws in wss_for_open]
        for ws_name in wss_name_for_open: 
            if log > 0: output.print_md("- :information: Будет открыт рн: **{}**".format(ws_name))
        workset_config.Open([ws.Id for ws in wss_for_open])
    else: workset_config = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
    output.print_md("- :information: Выбор типа открытия")
    options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets # отсоединять модель и сохранить рн
    output.print_md("- :information: Модель будет открыта **с отсоединением от ФХ и сохранением рабочих наборов**")

    options.Audit = audit  # активация проверки при открытие
    options.SetOpenWorksetsConfiguration(workset_config)
    o_timer = coreutils.Timer() 
    output.print_md("- :information: Настройки открытия сохранены**")


            #Фоновый режим
    doc = app.OpenDocumentFile(path, options)
    o_endtime = str(datetime.timedelta(seconds=o_timer.get_time())).split(".")[0]
    output.print_md("- :white_heavy_check_mark: Модель **{}** открыта в **фоне**. Время: **{}**".format(doc.Title,o_endtime))
    return doc


def closed_model(d):
    d.Close(False)
    output.print_md("-  :white_heavy_check_mark: **Модель RVT закрыта!**")


def get_project_path_from_ini(doc):
    """Получает путь ProjectPath из Revit.ini"""
    ini_file_path = os.path.join(str(doc.Application.CurrentUsersDataFolderPath), "Revit.ini")

    if not os.path.exists(ini_file_path):
        output.print_md("- :cross_mark: Файл Revit.ini не найден")
        return False

    try:
        with codecs.open(ini_file_path, 'r', encoding='utf-16') as file:
            for line in file:
                match = re.match(r'^\s*ProjectPath\s*=\s*(.*)', line)
                if match:
                    path = match.group(1).strip()
                    if os.path.exists(path):
                        return path
                    else:
                        print("Путь в ini недействителен: {}".format(path))
                        break  # выйдем и предложим выбор вручную

        # Если не нашли или путь недействителен — fallback
        fallback_path = r"D:\Revit 2021_Temp"
        print("Поиск в ini не дал результатов, пробую папку: {}".format(fallback_path))
        if os.path.exists(fallback_path):
            return fallback_path
        else:
            print("{} не валиден. Задайте папку вручную".format(fallback_path))
            path = forms.pick_folder(title="Выбор папки для сохранения локальных копий")
            if not path:
                print("Выбор отменен пользователем. Работа завершена")
                script.exit()
            return path

    except Exception as e:
        output.print_md("- :cross_mark: Ошибка открытия Revit.ini: **{}**".format(str(e)))
        return False


def select_file_in_txt(path_txt):
    # Чтение содержимого файла и запись в список lst_model_project
    lst_model_project = []
    try:
        with open(path_txt, 'r') as file:
            for line in file:
                lst_model_project.append(line.decode('utf-8').strip())
    except OSError as e:
        print("Ошибка при чтении файла {}: {}".format(path_txt, e))

    with forms.WarningBar(title="Выбор моделей"):
        items = forms.SelectFromList.show(lst_model_project,
                                            title='Выбор моделей',
                                            multiselect=True,
                                            button_name='Выбрать',
                                            width=800,
                                            height=800
                                            )
    if items: return items
    else: return None


#main

sel_models= None
name_view,open_wss,name_ws,save_html,open_folder,prefix,suffix = form_main()
if files_path:      sel_models = files_path
if files_txt:       sel_models = files_txt

projectpath = get_project_path_from_ini(doc)

if projectpath:
    if sel_models and folder_path:        
        name_html = "Отчет.html"
        output.print_md("##ЭКПОРТ NWC ({})".format(len(sel_models)))
        output.print_md("Папка для выгрузки: **{}**".format(folder_path))
        t_timer = coreutils.Timer()  
        output.update_progress(0, len(sel_models))

        for i, model_path in enumerate(sel_models):
            timer = coreutils.Timer()
            output.print_md("___") 
            output.print_md("###Модель: **{}**".format(model_path))
            targetPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(model_path)
            # docmodel= open_model(targetPath)

            docmodel= open_model( targetPath,
                activate = False,
                audit = False,
                detach = 1,
                closeallws=name_ws)
            
            ViewExportId = get_ViewExportId(docmodel,name_view)
            if not ViewExportId:
                closed_model(docmodel)
                continue
            try:
                export_NWC(docmodel,ViewExportId)
                closed_model(docmodel)
            except Exception as e:
                output.print_md("- :cross_mark: Ошибка экпорта! Код ошибки:" + str(e))                
                closed_model(docmodel)
                continue
            endtime = str(datetime.timedelta(seconds=timer.get_time())).split(".")[0]
            output.print_md("- :white_heavy_check_mark: **Работа с моделью завершена! Время: {} **".format(endtime))
            output.update_progress(i + 1, len(sel_models))

        t_endtime = str(datetime.timedelta(seconds=t_timer.get_time())).split(".")[0]
        output.print_md("___")   
        output.print_md("###**Завершено! Время: {} **".format(t_endtime))
        if save_html: output.save_contents(os.path.join(folder_path,name_html))
        if open_folder: coreutils.open_folder_in_explorer(folder_path)
    # else: print("НЕ ВЫБРАНЫ ФАЙЛЫ ДЛЯ ЭКСПОРТА ИЛИ ПАПКА ДЛЯ СОХРАНЕНИЯ")