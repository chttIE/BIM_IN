# -*- coding: utf-8 -*-
import codecs
import os
import re
from Autodesk.Revit.DB import   BuiltInParameter,  ImportPlacement, ModelPathUtils, OpenOptions, RevitLinkOptions,\
                                RevitLinkType, Transaction, Workset, WorksetConfiguration, WorksetConfigurationOption, \
                                RelinquishOptions, ReloadLatestOptions, \
                                SaveOptions, SynchronizeWithCentralOptions, \
                                TransactWithCentralOptions,SaveAsOptions,\
                                DetachFromCentralOption, WorksetKind,RevitLinkInstance,\
                                FilteredWorksetCollector as FWC,\
                                FilteredElementCollector as FEC,\
                                WorksetDefaultVisibilitySettings as WDVS, WorksharingUtils,\
                                TransmissionData, ExternalFileReferenceType


from pyrevit import forms, script,coreutils
import datetime

from sup import lst_len

output = script.get_output()
uiapp = __revit__
app = __revit__.Application
uidoc = __revit__.ActiveUIDocument
doc = __revit__.ActiveUIDocument.Document
user = app.Username

# --- 1) Утилита: выгрузить RVT-связи ДО открытия файла ---
def _unload_rvt_links_before_open(model_path, mark_transmitted=False, logfn=None):
    """
    Делает все RVT-связи файла незагружаемыми (Load=False) в TransmissionData.
    Работает только для локальных/сетевых файлов и когда файл закрыт.
    """
    tdata = TransmissionData.ReadTransmissionData(model_path)
    if tdata is None:
        if logfn: logfn(u"- ℹ️ Нет TransmissionData (нет внешних ссылок или облачный путь).")
        return 0

    changed = 0
    for xref_id in tdata.GetAllExternalFileReferenceIds():
        xref = tdata.GetLastSavedReferenceData(xref_id)
        if not xref:
            continue
        if xref.ExternalFileReferenceType == ExternalFileReferenceType.RevitLink:
            mp = xref.GetPath()
            ptype = xref.PathType
            # выключаем загрузку RVT-связи
            tdata.SetDesiredReferenceData(xref_id, mp, ptype, False)
            changed += 1
            if logfn:
                vis = ModelPathUtils.ConvertModelPathToUserVisiblePath(mp)
                logfn(u"- 🔌 RVT-связь будет открыта как выгруженная: `{}`".format(vis))

    if changed:
        if mark_transmitted:
            tdata.IsTransmitted = True
        TransmissionData.WriteTransmissionData(model_path, tdata)
    return changed



def get_ws_for_open(mp, name_ws):
    ws_for_open = []
    wss_link = WorksharingUtils.GetUserWorksetInfo(mp)
    for ws in wss_link:
        for name in name_ws:
            if "{}".format(name) in ws.Name and "Отвер" not in ws.Name:
                ws_for_open.append(ws)
    return ws_for_open

def convert_path(path):
    try:
        revitPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(path)
        if revitPath.IsValidObject: return revitPath
        else: print("НЕ ВАЛИДНАЯ ССЫЛКА")
    except:
        return False

def closed_model(d):
    d.Close(False)
    output.print_md("-  :white_heavy_check_mark: **Модель RVT закрыта!**")


def open_model( path,
                activate = True,
                audit = True,
                detach = 1,
                closeallws = True,
                log=2,
                unload='rvt',
                mark_transmitted=None):

    """
    Функция открытия модели
    
    Args:
        path - путь до модели <Autodesk.Revit.DB.ModelPath object at 0x0000000000003E0C [Autodesk.Revit.DB.ModelPath]>
        activate - bool. Способ открытия. В фоне или с активацией
            True - Обычное открытие с активацией модели
            False - Открытие в фоновом режиме
            
        audit - bool. Проверка  
            True - делать проверку.
            False - не делать проверку.

        detach - int. Варианты отсоединения модели
            0 - не отсоединять модель
            1 - отсоединять модель и сохранить рн
            2 - отсоединять модель и не сохранить рн
            3 - После открытия немедленно сохраните ее с текущим именем и снимите флаг передачи.
        
        closeallws - bool. Закрывать ли рабочие наборы
            True - Закрывать.
            False - Не закрывать.
            НЕ ДОПИСАЛ [list] - закрыть, но открыть те что что совпадают с именами в списке 
    
    Returns:
        (DB.Document): Документ текущей модели
    
    Examples:
        ```python
        from models import open_model
        open_model(path = targetPath,
                       activate=False, 
                       audit = 0,
                       detach = False, 
                       closeallws = True)
 
        <Autodesk.Revit.DB.Document object>,
        ```
    """
    # Задаем настройки открытия моделей
    
    if closeallws == True:
        workset_config = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        if log > 0: output.print_md("- :information: Рабочие наборы будут **закрыты**")
    elif closeallws == False: 

        workset_config = WorksetConfiguration(WorksetConfigurationOption.OpenAllWorksets)
        if log > 0: output.print_md("- :information: Рабочие наборы будут **открыты**")
    elif isinstance(closeallws,list):

        workset_config = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        wss_for_open = get_ws_for_open(path,closeallws)
        wss_name_for_open = [ws.Name for ws in wss_for_open]
        for ws_name in wss_name_for_open: 
            if log > 0: output.print_md("- :information: Будет открыт рн: **{}**".format(ws_name))
        workset_config.Open([ws.Id for ws in wss_for_open])
    else: workset_config = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
    
    options = OpenOptions()
    if detach == 0:
        options.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
        if log > 0: output.print_md("- Открытие **без отсоединения**")
    elif detach == 1:
        options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
        if log > 0: output.print_md("- Открытие **с отсоединением** (сохранить РН)")
    elif detach == 2:
        options.DetachFromCentralOption = DetachFromCentralOption.DetachAndDiscardWorksets
        if log > 0: output.print_md("- Открытие **с отсоединением** (не сохранять РН)")
    elif detach == 3:
        # Специальный режим: сохранить РН + выгрузить RVT-связи в TransmissionData
        options.DetachFromCentralOption = DetachFromCentralOption.DetachAndPreserveWorksets
        DetachFromCentralOption.ClearTransmittedSaveAsNewCentral
        if log > 0: output.print_md("- Открытие **с отсоединением** (сохранить РН) ")
        # try:
        #     # помечаем как переданный: опционально, но уместно в этом сценарии
        #     _unload_rvt_links_before_open(path, mark_transmitted=True,
        #                                   logfn=(output.print_md if log > 1 else None))
        # except Exception as ex:
        #     if log > 0:
        #         output.print_md(u"- ⚠️ Не удалось выгрузить RVT-связи через TransmissionData: **{}**".format(ex))
    else:
        # на всякий
        options.DetachFromCentralOption = DetachFromCentralOption.DoNotDetach
    options.Audit = audit  # активация проверки при открытие
    options.SetOpenWorksetsConfiguration(workset_config)
    
    o_timer = coreutils.Timer() 
    try:
        if activate:
            #Обычный режим
            uidoc = __revit__.OpenAndActivateDocument(path, options, False) # pyright: ignore[reportUndefinedVariable]
            o_endtime = str(datetime.timedelta(seconds=o_timer.get_time())).split(".")[0]
            if log > 1: output.print_md("- :white_heavy_check_mark: Модель **{}** открыта. Время: **{}**".format(uidoc.Document.Title,o_endtime)) 
            return uidoc       
        else:
            #Фоновый режим
            doc = app.OpenDocumentFile(path, options)
            o_endtime = str(datetime.timedelta(seconds=o_timer.get_time())).split(".")[0]
            if log > 1: output.print_md("- :white_heavy_check_mark: Модель **{}** открыта в **фоне**. Время: **{}**".format(doc.Title,o_endtime))
            return doc
    except Exception as ex:
        output.print_md("- :cross_mark: Ошибка открытия файла! Код ошибки:" + str(ex))

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



def Synchronize_models( d,
                        sinhr = True,
                        flatting = True,
                        relinquish_all = True,
                        reloadlatest = False,
                        save = True,
                        comm = True, 
                        comment = "BIM_IN Синхронизация",
                        log = 2):
    """
    log = 0 без принтов
    log = 1 принты основных моментов
    log = 2 все принты 
    """
    n_model = d.Title
    if log > 1: output.print_md("___")   
    if log > 1: output.print_md("###Модель: **{}**".format(n_model))
    if d.IsWorkshared:
        timer = coreutils.Timer()
        trans_options = TransactWithCentralOptions()
        sync_options = SynchronizeWithCentralOptions()
        if flatting: 
            sync_options.Compact = True 
            # if log > 0: output.print_md("- :information: **Модель будет сжата**")
        sync_options.SaveLocalAfter = True
        sync_options.SaveLocalBefore = True
        relinq_all = relinquish_all #Отдать все рн
        relinquish_options = RelinquishOptions(relinq_all)
        reload_latest_options = ReloadLatestOptions()
        save_options = SaveOptions()
        sync_options.SetRelinquishOptions(relinquish_options)
        if comm: 
            sync_options.Comment = comment
        if save:
            d.Save(save_options)
            output.print_md("-  :white_heavy_check_mark: **Сохранение модели**")
        if reloadlatest:
            d.ReloadLatest(reload_latest_options)
            if log > 0: output.print_md("-  :white_heavy_check_mark: **Обновление до последней версии**")
        if save:
            d.Save(save_options)
            if log > 0: output.print_md("-  :white_heavy_check_mark: **Сохранение модели**")
        if sinhr:
            d.SynchronizeWithCentral(trans_options , sync_options)
            if log > 0: output.print_md("-  :white_heavy_check_mark: **Синхронизация модели**")
        endtime = timer.get_time()
        endtime_hms = str(datetime.timedelta(seconds=endtime).seconds)
        if log > 0: output.print_md("- :white_heavy_check_mark: **Работа с моделью завершена. Время: {}с**".format(endtime_hms))
        return True
    else :
        if log > 0: output.print_md("- :information: Для документа не настроена совместная работа и синхронизация")



def select_file_local():
    #Функция не используется. 
    # folder_path = forms.pick_file(file_ext='txt', multi_file=False)
    folder_path = forms.pick_folder(title="Выберите папку куда сохранили результаты сканирования сервера")

    if not folder_path: script.exit()

    def list_files_in_folder(folder_path):
        lst_model = []
        try:
            for file in os.listdir(folder_path):
                lst_model.append(file.split(".txt")[0])
        except OSError as e:
            print("Ошибка чтения папки {}: {}".format(folder_path, e))
        return lst_model

    sel = list_files_in_folder(folder_path)

    if sel:
        selected_file = forms.SelectFromList.show(sel,
                                                title="Выбор объекта",
                                                width=400,
                                                button_name='Выбрать')
    if selected_file:
        file_path = os.path.join(folder_path, selected_file)
        
        # Чтение содержимого файла и запись в список lst_model_project
        lst_model_project = []
        try:
            with open(file_path+ ".txt", 'r') as file:
                for line in file:
                    lst_model_project.append(line.decode('utf-8').strip())
        except OSError as e:
            print("Ошибка при чтении файла {}: {}".format(file_path, e))

        with forms.WarningBar(title="Выбор моделей"):
            items = forms.SelectFromList.show(lst_model_project,
                                                title='Выбор моделей',
                                                multiselect=True,
                                                button_name='Выбрать',
                                                width=800,
                                                height=800
                                                )
        if items: return items
        else: script.exit()

def save_as_model(d,
                  path,
                  compact=True,
                  maxbackups=1,
                  overwrite=True):
    sop = SaveAsOptions()
    sop.Compact = compact
    sop.MaximumBackups = maxbackups
    sop.OverwriteExistingFile = overwrite

    try: 
        d.SaveAs(path, sop)
        output.print_md("-  :white_heavy_check_mark: **Модель сохранена**")
        return True
    except Exception as e:
        output.print_md("- :cross_mark: Ошибка сохранения модели **{0}**.\
                         Ошибка: **{1}**".format(d.Title, str(e)))
        return False



def set_workset_visibility(d, workset, boolean):
    WDVS.GetWorksetDefaultVisibilitySettings(d).SetWorksetVisibility(workset.Id, boolean)

def pin(el, status):
    el.get_Parameter(BuiltInParameter.ELEMENT_LOCKED_PARAM).Set(status)

def _owner(d,el):
    owner = WorksharingUtils.GetWorksharingTooltipInfo(d,el.Id).Owner
    if not owner or owner == str(__revit__.Application.Username):
        return False
    else: return owner

def create_ws_for_links(doc,log=1,pin_link=True):
    """
    Создет РН для связей, переносит связи в них и закрепляет
    
    Args:
        doc - документ <Autodesk.Revit.DB.Document object>
        log - логированией действий
        0 - отключает все уведомления
            1 - только Ballon уведомление
            2 - принты
        pin_link - bool. Закреплять ли связи
            True - Закреплять
            False - Не закреплять

    """
    pref = forms.ask_for_string(title="Префикс для РН",default="RVT-")
    if not pref: script.exit()
    # if forms.check_workshared():
    worksets = FWC(doc).OfKind(WorksetKind.UserWorkset).ToWorksets()
    names_of_worksets = {workset.Name for workset in worksets}

    with Transaction(doc, 'BIM_IN | Создание РН связей') as t:    
        t.Start()
        cnt=0
        links = FEC(doc).OfClass(RevitLinkInstance)
        if log == 2: output.print_md("- :information: Модуль создания рабочих наборов для связей")
        if log == 2: output.print_md(">> :information: Количество связей: **{}**".format(lst_len(links)))
        for link in links:
            
            if pin_link: pin(link,1)
            link_type = doc.GetElement(link.GetTypeId())
            if not link_type.IsNestedLink:
                param_inst = link.Parameter[BuiltInParameter.ELEM_PARTITION_PARAM]
                param_type = link_type.Parameter[BuiltInParameter.ELEM_PARTITION_PARAM]
                to_find = re.compile(".rvt|.RVT|.ifc|.IFC")
                match_obj = to_find.search(link.Name)
                the_index = match_obj.start()
                name = link.Name[:the_index]
                new_name = pref + name
                if new_name not in names_of_worksets:
                    try:
                        new_workset = Workset.Create(doc, new_name)
                        worksets.Add(new_workset)
                        names_of_worksets.add(new_workset.Name)
                        if log == 2: output.print_md(">> :white_heavy_check_mark: Рабочий набор: **{}** создан!".format(new_workset.Name))
                        cnt=+1
                    except Exception as e:
                        if log == 2: output.print_md(">> :cross_mark: Рабочий набор: **{}** не создан! Ошибка: **{}**".format(new_name,str(e)))
                for workset in worksets:
                    if new_name == workset.Name:
                        try:
                            own_inst = _owner(doc,link)
                            own_type = _owner(doc,link_type)
                            if not own_inst or not own_type:
                                param_inst.Set(workset.Id.IntegerValue)
                                param_type.Set(workset.Id.IntegerValue)
                            else: 
                                if log == 2: output.print_md(">> :cross_mark: Ошибка смены рн у связи. Элемент **{}** занят пользователем **{}-{}**".format(link.Name,own_inst,own_type))
                                continue
                        except Exception as e:
                            if log == 2: output.print_md(">> :cross_mark: Ошибка: **{}** при работе с рн {}".format(str(e),workset.Name))
                            continue
       
        if log == 1: forms.show_balloon("Рабочие наборы", "Рабочие наборы связей созданы. кол-во: {}".format(cnt))                
        if log == 2: output.print_md(">> :information: Создано новых рабочих наборов для связей: **{}**".format(cnt))
        t.Commit()



def create_local_model(model,projectpath):
  
    """
    Функция создания локальной копии модели
    
    Args:
        model - путь до исходной модели (строка).
        projectpath - путь до папки, где будет сохранена локальная копия (строка).
    
    Returns:
        (строка): Путь до созданной локальной копии модели или False в случае ошибки.
    
    Examples:
        ```python
        from models import createlocal
        target_path = createlocal(model="C:\\path\\to\\original_model.rvt", projectpath="C:\\path\\to\\project_folder")
        ```
    """
    c_timer = coreutils.Timer()
    folderforsave = os.path.normpath(projectpath)
    model_name = os.path.basename(model).split('.rvt')[0] 
    model_new_name = "{}_{}.rvt".format(model_name,user)
    new_path = os.path.join(folderforsave, model_new_name)
    output.print_md("- :information: Новый путь до локальной копии: **{}**".format(new_path))
    if os.path.exists(new_path):
        output.print_md("- :information: Файл **{}** существует. Будет удален".format(os.path.basename(new_path)))
        os.remove(new_path)
    try:
        targetPath = convert_path(new_path)
        modelpath = convert_path(model)
        WorksharingUtils.CreateNewLocal(modelpath, targetPath)
        c_endtime = str(datetime.timedelta(seconds=c_timer.get_time())).split(".")[0]
        output.print_md("- :white_heavy_check_mark: Локальная копия создана: **{}**. Время: **{}**".format(os.path.basename(new_path),c_endtime))    
    except Exception as ex:
        output.print_md("- :cross_mark: Ошибка при попытке создание локальной копию: {}".format(str(ex)))
        return False
    return targetPath


def add_link(d,path, placement_method=0, closed_ws=False, type_link=False):
    """
    Функция добавления ссылки в Revit
    
    Args:
        d: Document модели с которой ведется работа
        path (str): Путь до модели для добавления ссылки.
        placement_method (int): Способ размещения связи:
            0 - По общим координатам
            1 - В начале координат
            2 - По центру координат
            3 - В базовой точке  
        closed_ws (bool): Закрыть ли все рабочие наборы связей.
        type_link (bool): Тип связи:
            True - ссылка использует относительный путь. 
            False - используется абсолютный путь.
    
    Returns:
        RevitLinkInstance или bool: Экземпляр связи или False в случае ошибки.
    
    Examples:
        ```python
        from models import add_link
        add_link(doc = __revit__.ActiveUIDocument.Document,path = "C:\\path\\to\\link_model.rvt")
        ```
    """
    timer = coreutils.Timer()
    try:
        name_model = os.path.basename(path)
    except:
        name_model = path
    model_path = convert_path(path)
    if closed_ws:
        wsc = WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets)
        rlo = RevitLinkOptions(type_link, wsc)
    else:
        rlo = RevitLinkOptions(type_link)
    
# try:
    rl_type = RevitLinkType.Create(d, model_path, rlo)
    if placement_method == 0:
        method = [ImportPlacement.Shared, "По общим координатам"]
    elif placement_method == 1:
        method = [ImportPlacement.Origin, "В начале координат"]
    elif placement_method == 2:
        method = [ImportPlacement.Centered, "По центру координат"]
    elif placement_method == 3:
        method = [ImportPlacement.Site, "В базовой точке"]
    else:
        raise ValueError("Invalid placement_method value")
    
    try:
        rl_inst = RevitLinkInstance.Create(d, rl_type.ElementId, method[0])
        endtime = str(datetime.timedelta(seconds=timer.get_time())).split(".")[0]
        output.print_md("- :white_heavy_check_mark: Связь **{}** добавлена. **{}** {} Время: **{}**".format(
                                                    name_model, 
                                                    method[1], 
                                                    output.linkify(rl_inst.Id), 
                                                    endtime))
        return rl_type
    
    except:
        rl_inst = RevitLinkInstance.Create(d, rl_type.ElementId, ImportPlacement.Origin)
        endtime = str(datetime.timedelta(seconds=timer.get_time())).split(".")[0]
        output.print_md("- :white_heavy_check_mark: Связь **{}** добавлена **в начало координат.** {} Время: **{}**".format(
                                                    name_model, 
                                                    output.linkify(rl_inst.Id), 
                                                    endtime))
        return rl_type
    # except Exception as e:
    #     output.print_md("- :cross_mark: Ошибка создания ссылки: ({})".format(str(e)))
    #     return False