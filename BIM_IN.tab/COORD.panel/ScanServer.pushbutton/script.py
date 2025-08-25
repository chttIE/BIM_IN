# -*- coding: utf-8 -*-
import codecs
import os
import os.path as op
import urllib

from rpws import RevitServer, server as rs_mod 
from pyrevit import script,forms
output = script.get_output()
output.close_others(True)

# если нужной версии нет — добавим
def ensure_rs_version(ver):
    if ver not in rs_mod.sroots:
        # Для всех версий после 2012 шаблон одинаковый:
        # /RevitServerAdminRESTService{YYYY}/AdminRESTService.svc
        suffix = "" if ver == "2012" else ver
        rs_mod.sroots[ver] = "/RevitServerAdminRESTService{}/AdminRESTService.svc".format(suffix)

# пример: включаем 2019–2024 при необходимости
for v in ("2019","2020","2021","2022","2023","2024"):
    ensure_rs_version(v)

def _get_func(f):
    # достаём «голую» функцию из staticmethod, если нужно
    return getattr(f, '__func__', f)

def patch_rpws_api_path():
    if getattr(RevitServer, '_api_path_patched', False):
        return

    orig = _get_func(RevitServer._api_path)

    def _api_path_quoted(nodepath=None, _orig=orig, _urllib=urllib):
        apipath = _orig(nodepath)
        if isinstance(apipath, unicode):
            apipath = apipath.encode('utf-8')
        # _urllib теперь захвачено в замыкании и всегда доступно
        return _urllib.quote(apipath, safe='|/')

    RevitServer._api_path = staticmethod(_api_path_quoted)
    RevitServer._api_path_patched = True

# вызови один раз в начале скрипта:
patch_rpws_api_path()


def save_list_to_txt(data_list):
    """
    data_list: список строк, которые надо записать в txt
    """

    # --- Выбор папки ---
    folder_path = forms.pick_folder()
    if not folder_path:
        forms.alert("Папка не выбрана. Отмена.")
        return

    # --- Запрос имени файла ---
    file_name = forms.ask_for_string(
        prompt="Введите имя объекта:",
        default="ОБЪЕКТ1"
    )
    if not file_name:
        forms.alert("Имя файла не задано. Отмена.")
        return

    file_path = os.path.join(folder_path, file_name + ".txt")

    # --- Запись списка ---
    try:
        with codecs.open(file_path, "w", encoding="utf-8") as f:
            for line in data_list:
                f.write(str(line) + "\n")

        forms.alert("Файл успешно создан:\n{}".format(file_path))
    except Exception as e:
        forms.alert("Ошибка при записи файла: {}".format(e))

# --- Дальше обычная работа с сервером ---


HOST = forms.ask_for_string(prompt="Введите IP адрес сервера",title="Сканер")
VER = forms.ask_for_string(prompt="Введите версию RevitServer",title="Сканер",default="2022")

if not HOST or not VER: script.exit() 

try:
    rs = RevitServer(HOST, VER)
except:
    print("Указанная версия ревита пока недоступа")
    
def to_rsn(host, rel_path):
    # в rsn:// используем прямые слэши
    return u"rsn://{}/{}".format(host, rel_path.lstrip(u"\\").replace(u"\\", u"/"))

# Корневые папки (ВАЖНО: без аргумента или op.sep, а не "/")
root_folders = rs.listfolders()  # = rs.listfolders(op.sep)

lst_rsn = []

for parent, folders, files, models in rs.walk(top=None, topdown=True, digmodels=False):
    if models:
        # output.print_md(u"### 📂 Папка `{}` — моделей: **{}**".format(parent or u"\\", len(models)))
        for m in models:
            minfo = rs.getmodelinfo(m.path)
            rsn = to_rsn(rs.name, m.path)
            lst_rsn.append(rsn)
            print(u"{} - {} - {} - {} - {}".format(rsn, minfo.size,minfo.date_created,minfo.date_modified,minfo.last_modified_by))

if lst_rsn: save_list_to_txt(lst_rsn)