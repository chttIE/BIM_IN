# -*- coding: utf-8 -*-

from Autodesk.Revit.DB import FilteredElementCollector as FEC
import math
import os
from pyrevit import script
from pyrevit import output
from Autodesk.Revit import DB
from Autodesk.Revit.DB import *
from rpw import DB

doc = __revit__.ActiveUIDocument.Document  # noqa
linked_doc = doc
stylesheet = r"Z:\Bim\13. Прочее\pyRevit\outputstyles.css"
output.set_stylesheet(stylesheet)
output = script.get_output()
lfy = output.linkify
script.get_output().close_others(all_open_outputs=True)
output.set_title("Список изображений")
output.set_width(1500)
output.set_font('Atyp Text', 12)


def get_ws_el(el):
    try:
        return el.get_Parameter(DB.BuiltInParameter.ELEM_PARTITION_PARAM).AsValueString()
    except:
        return "Error"

def get_creator(el,linked_doc):
    try:
        return DB.WorksharingUtils.GetWorksharingTooltipInfo(linked_doc,el.Id).Creator
    except:
        return "-"

def get_size(img):
    if os.path.exists(img.Path):
        file_size_bytes = os.path.getsize(img.Path)
        if file_size_bytes == 0:
            return "0B"
        size_unit = ("B", "KB", "MB", "GB", "TB", "PB", "EB", "ZB", "YB")
        i = int(math.floor(math.log(file_size_bytes, 1024)))
        p = math.pow(1024, i)
        size = round(file_size_bytes / p, 2)
        return "{}{}".format(size, size_unit[i])
    else:
        return "-"

def get_name_img(img):
    try:
        return os.path.basename(img.Path)
    except:
        return "-"
    
def lst_len(lst):
    return sum(1 for _ in lst)

def get_img(linked_doc, _):
    imgs = FEC(linked_doc)\
              .OfClass(ImageType)\
              .ToElements()
    cnt_img=lst_len(imgs)
    if cnt_img != 0:
        output.print_md("___")
        output.print_md("## ПРОВЕРКА ИЗОБРАЖЕНИЙ")
        img_data = sorted([(    get_name_img(img),
                                lfy(img.Id),
                                get_creator(img,linked_doc),
                                get_ws_el(img),
                                img.Path,
                                img.PathType,
                                img.Source,
                                img.Status,
                                str(img.WidthInPixels)+"х"+str(img.HeightInPixels),
                                get_size(img) )
                                for img in imgs], 
                                key=lambda x: (x[5]))
        img_data = [(i + 1,) + x for i, x in enumerate(img_data)]
        output.print_table( table_data=img_data,
                            title="**Изображений** (кол-во: {})".format(cnt_img),
                            columns=["№","Имя", "ID","Автор","РН", "Загруженно из файла","Тип пути","Тип связи", "Статус связи","Размер","Вес"],
                            formats=['', '','', '', '','', '', '', '', ''])
    else:
        output.print_md("___")
        output.print_md("## В МОДЕЛЕ ОТСУТСТВУЮТ ИЗОБРАЖЕНИЙ")  


# main
get_img(linked_doc,1)