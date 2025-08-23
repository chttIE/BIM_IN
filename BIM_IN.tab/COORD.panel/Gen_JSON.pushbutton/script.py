# -*- coding: utf-8 -*-
__title__ = 'Сканирование серверов'
__author__ = 'IliaNistratov'
__doc__ =   """
            Генерит json файл с ссылками на модели.
            """
import os
import json
import codecs
import datetime
import System.Diagnostics as Diagnostics # type: ignore
from pyrevit import script, forms, coreutils
from sup import get_size_file
output = script.get_output()
script.get_output().close_others(all_open_outputs=True)
output.set_width(900) 

#Путь до файла txt где записаны адреса до объектов

try:
    doc = __revit__.ActiveUIDocument.Document  # type: ignore
    version = doc.Application.VersionNumber
except:
    version = "2021"


import os
import json
import codecs
from pyrevit import forms

def get_config_path_from_file_or_folder(config_filename='settings.json', key='path', is_folder=False):
    script_dir = os.path.dirname(__file__)
    config_path = os.path.join(script_dir, config_filename)

    config = {}

    # Попытка загрузить конфиг
    if os.path.exists(config_path):
        try:
            with codecs.open(config_path, 'r', encoding='utf-8') as f:
                config = json.load(f)
        except:
            pass
    else:
        forms.alert("Настройки отсутствует!\n\nЗадайте настройки")       
    # Проверка на наличие и валидность пути
    if key in config:
        saved_path = config[key]
        if is_folder:
            if not os.path.isdir(saved_path):
                forms.alert("Сохранённая папка недоступна:\n{}".format(saved_path))
                config[key] = None
        else:
            if not os.path.isfile(saved_path):
                forms.alert("Сохранённый файл недоступен:\n{}".format(saved_path))
                config[key] = None

    # Если путь валиден — возвращаем
    if key in config and config[key]:
        return config[key]

    # Запрос пути
    if is_folder:
        selected = forms.pick_folder(title="Выберите папку куда сохранять рузультаты сканирования")
    else:
        selected = forms.pick_file(file_ext='txt', title="Укажите .txt с адресами")

    if not selected or (is_folder and not os.path.isdir(selected)) or (not is_folder and not os.path.isfile(selected)):
        forms.alert("Выбранный путь недействителен. Операция прервана.", exitscript=True)

    # Сохраняем в конфиг
    config[key] = selected
    try:
        with codecs.open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4, ensure_ascii=False)
    except Exception as e:
        forms.alert("Ошибка при сохранении настроек:\n{}".format(str(e)))

    return selected




def get_folder_size(folder_path):
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(folder_path):
        for filename in filenames:
            file_path = os.path.join(dirpath, filename)
            total_size += os.path.getsize(file_path)
    return total_size / (1024 * 1024)

def get_base_lst(path):
    base_lst = []
    if os.path.exists(path):
        with open(path, 'r') as file:
            for line in file:
                base_lst.append(line.decode('utf-8').strip())
    else:
        
        print("Путь {} не существует.".format(path))
    return base_lst

def open_file(file_path):
    try:
        Diagnostics.Process.Start(file_path)
    except Exception as e:
        print("Произошла ошибка при попытке открыть файл:", e)

def replace_link(link):
    clearlink = link.replace('\Revit Server 2022',"")
    clearlink = clearlink.replace('\Revit Server 2021\Projects',"")
    clearlink = clearlink.replace('\Revit Server 2019\Projects',"")
    clearlink = clearlink.replace('\Revit Server 2020\Projects',"")
    clearlink = clearlink.replace('\RevitServer2021\Projects',"")
    clearlink = clearlink.replace('\Revit Server 2024',"")
    clearlink = clearlink.replace('\Revit Server',"")
    clearlink = clearlink.replace('\Projects',"")
    return "rsn:" + clearlink

import os

def find_models(path, extension=".rvt"):
    json_data_lst = []
    lst_exception = ["!Архив", "Архив"]

    # Проверяем, является ли путь сетевым
    if str(path).startswith("\\") and "RevitShare" not in str(path):
        for root, dirs, files in os.walk(path):
            for dir_name in dirs:
                if dir_name.endswith(extension):  # Проверка на правильный объект (должен быть file, а не dir)
                    folder_path = os.path.join(root, dir_name)
                    parent_path = os.path.dirname(folder_path)
                    m_disc = os.path.basename(parent_path)

                    if m_disc not in lst_exception:
                        # Оптимизация с использованием множества set
                        base_files = {"02. Координация", "01. Базовые файлы", "01.Базовые файлы", "00.Базовые файлы", "00. Базовые файлы", "01. Базовые"}
                        links = {"02. Связи", "02.Связи"}

                        m_disc = "BF" if m_disc in base_files else m_disc
                        m_disc = "LINK" if m_disc in links else m_disc

                        if len(m_disc) <= 8:
                            m_name = os.path.basename(folder_path).split(".rvt")[0]
                            m_size = get_folder_size(folder_path) if chk_size else "0"
                            m_path = replace_link(folder_path)
                            output.print_md("-  Раздел **{}**: Имя **{}**: Размер **{}** Мб Версия: **{}**".format(
                                m_disc,
                                m_name,
                                str(m_size) + " :information:" if m_size > 1000 else m_size,
                                m_version
                            ))

                            # Создание словаря с данными модели
                            m_data = {
                                "Address": m_path,
                                "Discipline": m_disc,
                                "Name": m_name,
                                "Size": m_size,
                                "Version": m_version
                            }

                            # Добавление словаря в результат
                            json_data_lst.append(m_data)
                        else:
                            m_name = os.path.basename(folder_path).split(".rvt")[0]
                            output.print_md("- :information: Модель **{}** в разделе **{}** найдена, но **не будет** добавлена в базу".format(m_name, m_disc))
    else:
        for root, dirs, files in os.walk(path):
            for file in files:
                if file.endswith(extension):
                    folder_path = os.path.join(root, file)  # Исправление: должно быть file вместо dir_name
                    m_name = os.path.basename(folder_path).split(".rvt")[0]
                    m_size = get_size_file(folder_path)
                    m_path = folder_path  # Исправление: должно быть folder_path вместо replace_link
                    output.print_md("-  Раздел **{}**: Имя **{}**: Размер **{}** Мб".format(
                        "N/A",  # m_disc не определено для не сетевых путей, поэтому используем "N/A"
                        m_name,
                        str(m_size) + " :information:" if m_size > 1000 else m_size
                    ))

                    # Создание словаря с данными модели
                    m_data = {
                        "Address": m_path,
                        "Discipline": "N/A",  # Здесь дисциплина неизвестна
                        "Name": m_name,
                        "Size": m_size
                    }

                    # Добавление словаря в результат
                    json_data_lst.append(m_data)

    return json_data_lst

def json_creation(json_name,json_data_lst):
    # Преобразование списка словарей в строку JSON с отступами
    json_data = json.dumps(json_data_lst, indent=4, ensure_ascii=False)
    # Запись JSON-строки в указанный файл с определенной кодировкой
    
    json_path = os.path.join(json_folder, json_name)
    with codecs.open(json_path, "w", "utf-8") as file:
        file.write(json_data)
    output.print_md(" :white_heavy_check_mark: Данные занесены в файл: **{}**".format(json_path))  

chk_size = False
path_model = get_config_path_from_file_or_folder(key='path_input', is_folder=False)
json_folder = get_config_path_from_file_or_folder(key='path_output_dir', is_folder=True)
#main
if __shiftclick__: #используем клик с шифтом что бы открыть файл хранящий список объектов
    open_file(path_model)
else:
    base_lst = get_base_lst(path_model) #функция для чтения txt и формирования списка
    if base_lst: #Проверяем не пустой ли список и даем выбрать элемент из него
        selected_file = forms.SelectFromList.show(  base_lst,
                                                    title="Выбор объекта",
                                                    width=800,
                                                    hight=700,
                                                    multiselect=True,
                                                    button_name='Выбрать')
    if selected_file: #Проверяем что произошел выбор и используем его
        if forms.alert("Проверять размер?\n\nЗанимает больше времени", yes=True,no=True):
            chk_size = True
        t_timer = coreutils.Timer()
        
        m_version = forms.ask_for_string(prompt="Подтвердите версию",title="Сканер",default=version)
        for i, sel in enumerate(selected_file): 
            output.print_md("##ОБЪЕКТ: **{}**".format(os.path.basename(sel)))
            output.print_md("___")           
            try:
                data = find_models(sel) #Ищем по выбранному пути модели и собираем их в список
                if data: #Проверяем что нашли модели
                    output.print_md(" :white_heavy_check_mark: Найдено моделей: **{}**".format(len(data)))
                    name_json = os.path.basename(sel)
                    name_json = forms.ask_for_string(prompt="Задайте имя для JSON файла",title="Сканер",default=name_json)
                    json_creation("{}.json".format(name_json),data)
                    output.print_md("___")  
                else:
                    output.print_md(" :information: Модели не найдены")
                    output.print_md("___")  
            except Exception as e:
                output.print_md("- :cross_mark: Ошибка с объектом {}. Ошибка: {}".format(sel,str(e)))
                output.print_md("___")
            output.update_progress(i + 1, len(selected_file))   
        t_endtime = str(datetime.timedelta(seconds=t_timer.get_time())).split(".")[0]
        output.print_md("**Время: {} **".format(t_endtime))