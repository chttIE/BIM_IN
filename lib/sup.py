
# -*- coding: utf-8 -*-
import ctypes
from Autodesk.Revit.DB import  Category, Element, ExternalResourceType, Family, MEPSystemType,\
      FilteredElementCollector as FEC, RevitLinkInstance, WorksharingUtils,Transaction,ElementId,BuiltInCategory
from pyrevit import forms, script
import os
import io

from pyrevit.coreutils import Guid
from pyrevit.framework import Diagnostics
# Импортируем .NET List
import clr
clr.AddReference('System')
from System.Collections.Generic import List

DEFAULT_INPUTWINDOW_WIDTH = 500

from Autodesk.Revit.DB import WorksharingUtils

class WorkSharingInfo:
    def __init__(self, doc):
        """
        Инициализация класса WorkSharingInfo для работы с рабочими наборами в документе Revit.
        :param doc: объект документа Revit
        """
        self.doc = doc

        # Проверка, является ли документ рабочим набором
        if not self.doc.IsWorkshared:
            raise Exception("У документа не настроена совместная работа")

    def get_LastChangedBy(self, el):
        """
        Возвращает информацию о последнем изменившем элемент.
        :param el: элемент Revit
        :return: строка с именем последнего изменившего или None если не удается получить информацию.
        """
        try:
            # Получаем информацию о последнем изменении элемента через WorksharingUtils
            wsti = WorksharingUtils.GetWorksharingTooltipInfo(self.doc, el.Id)
            return wsti.LastChangedBy
        except Exception as e:
            return None

    def get_Owner(self, el):
        """
        Возвращает создателя элемента.
        :param el: элемент Revit
        :return: Имя пользователя владеющего элемента или None если не удается получить информацию или элемент свободен.
        """
        try:
            wti = WorksharingUtils.GetWorksharingTooltipInfo(self.doc, el.Id) 
            own = wti.Owner
            if not own: return False
            else: 
                if str(own) == str(__revit__.Application.Username):
                    return False
            return wti.Owner
        except Exception as e:
            return None

    def get_Сreator(self, el):
        """
        Возвращает имя создателя элемента.
        :param el: элемент Revit
        :return: строка с именем создателя или None если не удается получить информацию.
        """
        try:
            wti = WorksharingUtils.GetWorksharingTooltipInfo(self.doc, el.Id)
            return wti.Creator
        except Exception as e:
            return None

    def get_Workset_info(self, el):
        """
        Возвращает всю доступную информацию о рабочем наборе элемента.
        :param el: элемент Revit
        :return: словарь с информацией о рабочем наборе или '-'.
        """
        try:
            wti = WorksharingUtils.GetWorksharingTooltipInfo(self.doc, el.Id)
            return {
                "Владелец": wti.Owner,
                "Последние изменения": wti.LastChangedBy,
                "Создатель": wti.Creator
            }
        except Exception as e:
            return None
        


def isNullOrWhiteSpace(str):
    if (str is None) or (str == "") or (str.isspace()):
        return True
    return False


def getpath_RevitLinkType(rlt):
    KEY = "5e6433a2-9679-4d1f-943e-c5215f772b8a"
    ert = ExternalResourceType(Guid(KEY)) 
    err = rlt.GetExternalResourceReferences()[ert]
    return err.GetReferenceInformation()["Path"] 


# # функция конвертации из внутренних едениц в милиметры
# def IUtoMM(value):
#     from Autodesk.Revit.DB import UnitUtils,DisplayUnitType
#     return UnitUtils.ConvertFromInternalUnits(value, DisplayUnitType.DUT_MILLIMETERS) 

# def toIU(value):
#     from Autodesk.Revit.DB import UnitUtils,DisplayUnitType
#     return UnitUtils.ConvertToInternalUnits(value, DisplayUnitType.DUT_MILLIMETERS)

# # функция конвертации из внутренних едениц в метры
# def IUtoM(value):
#     from Autodesk.Revit.DB import UnitUtils,DisplayUnitType
#     return UnitUtils.ConvertFromInternalUnits(value, DisplayUnitType.DUT_METERS)

# # функция конвертации из значения в еденицы типа параметра
# def INT_to_ParamUT(value, parameter):
#     from Autodesk.Revit.DB import UnitUtils
#     return UnitUtils.ConvertToInternalUnits(value, parameter.DisplayUnitType)


def open_file(file_path):
    try:
        Diagnostics.Process.Start(file_path)
    except Exception as e:
        print("Произошла ошибка при попытке открыть файл:", e)


def get_size_file(file_path):
    """
    Возвращает размер файла по пути в МБ
    """
    file_size = os.path.getsize(file_path)
    file_size_mb = file_size / (1024 * 1024) 
    
    return "{}".format(round(file_size_mb, 2))
    

def set_window_topmost(window_title):
    """
    Функция для закрепления окна поверх других.
    в качестве аргумента передаем "Имя окна" (str)
    """
    hwnd = ctypes.windll.user32.FindWindowW(None, window_title)
    if hwnd:
        ctypes.windll.user32.SetWindowPos(hwnd, -1, 0, 0, 0, 0, 0x0001 | 0x0002)
    else:
        script.get_logger().warning("Window '{}' not found.".format(window_title))

def lst_len(lst):
    """
    Функция для определния длины списка.
    """
    return sum(1 for _ in lst)



def get_lookup_param_el(el,name):
    try: return el.LookupParameter(name)
    except: return False





def read_json_obj(path):
    import json
    # Открытие и чтение файла JSON
    """Чтение JSON и возврат списка адресов"""
    with io.open(path, 'r', encoding='utf-8') as file:
        data = json.load(file)

    # Извлечение всех значений поля "Address"
    addresses = [entry['Address'] for entry in data]

    # Возвращаем список адресов
    return addresses





def get_existing_elements(doc,view):
    #Функцию надо переписать и избавиться от класса CategoryOption
    #Думаю в нем нет смысла так как это просто обертка что бы имена в методе SelectFromList выводить
    #А это можно провернуть через атреббут name_attr
    #Импорт вынес только сюда, что бы он просиходил только если используюется функция
    from cats import lst_cats
    """
    Данная фукнция позволяет сначала запросить у пользователя какая(ие) категория(ии) 
    элементов СУЩЕСТВУЮЩИХ в проекте ему интересна(ы), после выбора собирается и 
    возвращается список элементов
    
    Args:
        doc - активный документ <Autodesk.Revit.DB.Document object at 0x0000000000003E0C [Autodesk.Revit.DB.Document]>
        view - Вид на котором искать элементы <Autodesk.Revit.DB.View3D object at 0x0000000000003A81 [Autodesk.Revit.DB.View3D]>
            Если передать False поиск ведется во всей модели
    
    Returns:
        (list[DB.Element]): Лист элементов
    
    Examples:
        ```python
        from sup import get_existing_elements
        get_existing_elements(doc,False)
        [<Autodesk.Revit.DB.Mechanical.Duct object>,
        <Autodesk.Revit.DB.Mechanical.Duct object>] 
        ```
    """

    class CategoryOption:
        def __init__(self, name, value):
            self.name = name
            self.value = value

    def createCheckBoxe(catSet):
        with forms.WarningBar(title='Выбор категории элементов'):
            categoties_options = [CategoryOption(c.Name, c.Id) for c in catSet]
            if categoties_options != 0:
                catCheckBoxes = forms.SelectFromList.show(categoties_options,
                                                        multiselect=True,
                                                        title='Выберите категорию',
                                                        width=500,
                                                        button_name='Выбрать')
                return catCheckBoxes
            else: script.exit()


    elemList=[]
    #Сборка коллектора 
    category_name_set = set()
    for currentCat in lst_cats():
        frstElem = FEC(doc).OfCategory(currentCat).WhereElementIsNotElementType().FirstElement()
        if frstElem:
            category_name_set.add(frstElem.Category)

    check_box = createCheckBoxe(category_name_set)
    if check_box:
        selCatIds = [p.value for p in check_box]
        if view:
            for currId in selCatIds:
                elemList.extend(FEC(doc,view.Id).OfCategoryId(currId).WhereElementIsNotElementType().ToElements())
        else:
            for currId in selCatIds:
                elemList.extend(FEC(doc).OfCategoryId(currId).WhereElementIsNotElementType().ToElements())
    else: script.exit()
    return elemList

def get_RevitLinkInstance(doc,isload=True):
    """
    Возвращает экземпляры связей
    По дефолту показывает только загруженные связи
    
    Args:
        doc - документ <Autodesk.Revit.DB.Document object> 
        isload (bool) - Показывать только загруженные связи или все
    Returns:
        (list): Список связей [<Autodesk.Revit.DB.Document.RevitLinkInstance object>]
    Examples:
        ```python
        from models import get_RevitLinkInstance
        rli = get_RevitLinkInstance(doc)

    """
    links = FEC(doc).OfClass(RevitLinkInstance)
    if not links: forms.alert(msg="В моделе отсутствуют связи", title="Связи",sub_msg="",exitscript=True)
    if isload: 
        links = [link for link in links if link.GetLinkDocument()]
        if not links: forms.alert(msg="В моделе отсутствуют загруженные связи", title="Связи",sub_msg="",exitscript=True)
    
    sel_links = forms.SelectFromList.show(
        links,
        multiselect=True,
        name_attr = "Name",
        title='Выберите связи',
        width=600,
        button_name='Выбрать')
    if not sel_links:script.exit()
    return sel_links

def select_var(doc):
    #ПЕРЕПИСАТЬ
    """

    Дает выбрать с каким типом моделей работать
    1- связанные модели
    2- открытые модели
    """
    res = forms.alert("",
                    options=[
                            "Связанные модели",
                            "Открытые модели"])
    
    if res == "Связанные модели":
        links = get_RevitLinkInstance(doc)

        return [link.GetLinkDocument() for link in links]
    elif res == "Открытые модели":return forms.select_open_docs(title='Выбор документов',check_more_than_one=False)
    else:script.exit()

def get_name(el):
    return Element.Name.GetValue(el)

def sel_MEPSystem(  doc,
                    multiselect=False,
                    title = "Выбор системы",
                    filter = False):
    """
    Описание:
        Дает выбрать доступные в проекте системы.
    args:
        multiselect - велючает чекбоксы для выбора нескольких элементов
        filter - фильтр систем, следует передать список
        из типов систем которые необходимо отобразить
        Если False то отобразится все. 
        Допустипые:
                    [PipingSystemType,MechanicalSystemType]
        Возвращает список систем или одну выбранную систему 
    """


    mepsystem = sorted(FEC(doc).WhereElementIsElementType().OfClass(MEPSystemType), key=lambda x: x.Category)
    if filter and isinstance(filter, list):
        mepsystem_new =[]
        for m in mepsystem:
            for f in filter:
                if isinstance(m,f):        
                    mepsystem_new.append(m)    
        mepsystem = mepsystem_new   
    mepsystem_name = [get_name(m) for m in mepsystem]
    sel_mep = forms.SelectFromList.show(    mepsystem_name ,
                                            title=title,
                                            multiselect=multiselect)
    if not sel_mep: script.exit() 
    if multiselect:
        result=[]
        for m in mepsystem:
            for s in sel_mep:
                if Element.Name.GetValue(m) == s:
                    result.append(m)
        return result
    else:   
        for m in mepsystem:
            if Element.Name.GetValue(m) == sel_mep:
                    return m



#Проверка есть ли у категории параметр
def CheckCategoryInParameter(doc,parameterName, ost_cat):
    """
    Вернет True если к категории добавлен параметр

    Необходимо передать:
        doс - документ модели
        parameterName - имя параметра в виде строки
        ost_cat - BuiltInCategory.OST_*
    """
    
    category = Category.GetCategory(doc,ost_cat)
    BindingMap = doc.ParameterBindings
    iterator = BindingMap.ForwardIterator()
    oldCatSet = None

    while iterator.MoveNext():
        if iterator.Key.Name.Equals(parameterName):
            binding = iterator.Current
            oldCatSet = binding.Categories


    if oldCatSet is None:
        return False

    if oldCatSet.Contains(category):
        return True
    else:
        return False

def name_family(el,doc):
    try:
        family_symbol = doc.GetElement(el.GetTypeId())
        return family_symbol.FamilyName
    except:
        return "-"

def sel_open_fam(title='Select Open Familys',
                     button_name='OK',
                     width=DEFAULT_INPUTWINDOW_WIDTH,    #pylint: disable=W0613
                     multiple=True,
                     check_more_than_one=True,
                     filterfunc=None):
    """Standard form for selecting open documents.

    Args:
        title (str, optional): list window title
        button_name (str, optional): list window button caption
        width (int, optional): width of list window
        multiple (bool, optional):
            allow multi-selection (uses check boxes). defaults to True
        check_more_than_one (bool, optional): 
        filterfunc (function):
            filter function to be applied to context items.

    Returns:
        (list[DB.Document]): list of selected documents

    Examples:
        ```python
        from pyrevit import forms
        forms.select_open_docs()
        [<Autodesk.Revit.DB.Document object>,
         <Autodesk.Revit.DB.Document object>]
        ```
    """
    # find open documents other than the active doc
    app = __revit__.Application # type: ignore
    docs =  app.Documents
    open_docs_fam = [d for d in docs if d.IsFamilyDocument]    #pylint: disable=E1101
    # if check_more_than_one:
    #     open_docs.remove(revit.doc)    #pylint: disable=E1101

    if not open_docs_fam:
        forms.alert('Открытых семейств нет')
        return
    if open_docs_fam.Count == 1:
        return open_docs_fam
    else:
        return forms.SelectFromList.show(
            open_docs_fam,
            name_attr='Title',
            multiselect=multiple,
            title=title,
            button_name=button_name,
            filterfunc=filterfunc
            )


def collect_elements_on_view(doc,view=None,
                             exclude_categories=None,
                             exclude_classes=None,
                             preview='off'):
    """
    Собирает элементы, видимые на (view|ActiveView), с расширяемыми исключениями.
    Может включить временное превью:
      preview='off'        — без превью (по умолчанию)
      preview='isolate'    — временно изолировать собранные элементы
      preview='hide_others'— синоним isolate (скрыть всё остальное)
      preview='hide'       — временно скрыть САМИ собранные элементы

    :param view: View | None
    :param exclude_categories: iterable[BuiltInCategory]
    :param exclude_classes: iterable[type] (классы API, напр. CurveElement)
    :return: list[Element]
    """
    if view is None:
        view = __revit__.ActiveUIDocument.ActiveView
    category = []
    # дефолтные исключения: линии и камеры
    exclude_category_names=["Камеры", "Оси", "Линии моделей","Виды","Границы 3D вида"]
    col = (FEC(doc, view.Id)
           .WhereElementIsNotElementType())
    namecatforignore = ["Камеры","Оси"]
    result = []
    for el in col:

        # исключения по категории
        cat = el.Category
        try:
            cat_name = cat.Name
        except:
            cat_name = "ERROR"

        if cat is None:
            continue
        try:
            # BuiltInCategory у системных категорий — отрицательный int id
            bic = BuiltInCategory(cat.Id.IntegerValue)

            if cat_name not in category and cat_name not in exclude_category_names:
                category.append(cat.Name)
        except:
            # не BuiltInCategory — пропускаем проверку
            pass

        result.append(el)

    # превью через временную изоляцию/скрытие
    if preview and preview.lower() != 'off':
        _apply_temp_preview(doc,view, result, mode=preview.lower())

    return result,category


def _apply_temp_preview(doc,view, elements, mode='isolate'):
    """
    Включает временное превью на виде:
      - 'isolate' / 'hide_others' — IsolateElementsTemporary(IDs)
      - 'hide'                    — HideElementsTemporary(IDs)
    """
    ids = List[ElementId]([e.Id for e in elements])

    with Transaction(doc, "Preview collected elements") as t:
        t.Start()
        if mode in ('isolate', 'hide_others'):
            view.IsolateElementsTemporary(ids)
        elif mode == 'hide':
            view.HideElementsTemporary(ids)
        else:
            t.RollBack()
            return
        t.Commit()

def search_and_change_parameters_for_element(doc,elements,name_parameter,text):
    bad= []
    g,i = 0,0
    with Transaction(doc,"Заполнение параметра") as t:
        t.Start()
        try:
            for el in elements:
                p_d = el.LookupParameter(name_parameter)
                if p_d and not p_d.IsReadOnly: 
                    if p_d.AsString() == text:
                        i+=1
                        continue
                    else:
                        p_d.Set(text)
                        g+=1
                else:
                    bad.append(el)
            doc.Regenerate()
        except Exception as e:
            print("- ОШИБКА {} Нет параметра у {} - категория {}".format(str(e),el.Id,el.Category))

        t.Commit()
    return bad,g,i


def get_familysymbol(doc):
    """
    Данная функция вернет пользователю типоразмер выбранного семейства 
    
    Args:
        doc - активный документ <Autodesk.Revit.DB.Document object at 0x0000000000003E0C [Autodesk.Revit.DB.Document]>
    Returns:
        (list[DB.FamilySymbol]): Типоразмер семейства
    
    Examples:
        ```python
        from sup import sel_familysymbol
        sel_familysymbol(doc)
        <Autodesk.Revit.DB.FamilySymbol object>,
        ```
    """


    family_coll = FEC(doc).OfClass(Family)
    sel_family = forms.SelectFromList.show(family_coll,
                                                    title="Выбор семейства",
                                                    name_attr='Name',
                                                    multiselect=False,)
    if sel_family:
        familysymbolids = sel_family.GetFamilySymbolIds()
        familysymbol = [doc.GetElement(familysymbolid) for familysymbolid in familysymbolids]
        sel_familysymbol = forms.SelectFromList.show([Element.Name.GetValue(f) for f in familysymbol],
                                                title="Выбор типоразмера",
                                                multiselect=False,)
        if sel_familysymbol:
            for f in familysymbol:
                if Element.Name.GetValue(f) == sel_familysymbol:
                    return f
        else: script.exit()
    else: script.exit()