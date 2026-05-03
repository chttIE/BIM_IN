# -*- coding: utf-8 -*-
__title__ = 'Открыть семейства'
__doc__ = ""
__highlight__ = 'new'
__context__ = "zero-doc"

import os
from Autodesk.Revit.DB import (Document, ExportRange,  FitDirectionType, ImageExportOptions, ImageFileType, ImageResolution, ZoomFitType)
from pyrevit import revit, script


uiapp = __revit__ # type: ignore
uidoc = __revit__.ActiveUIDocument  # type: ignore
app = __revit__.Application          # type: ignore
active_doc = uidoc.Document          # type: ignore

output = script.get_output()
script.get_output().close_others(all_open_outputs=True)


# --- utils ---
def export_image(cur_doc, folder_for_save):
    """
    Экспорт активного вида в PNG (600 DPI) в подпапку "_image".
    """
    # output.print_md("- Экспорт миниатюры")
    if cur_doc is None or not isinstance(cur_doc, Document):
        raise ValueError("cur_doc должен быть объектом Revit Document")

    active_view = cur_doc.ActiveView
    if active_view is None:
        return None

    # создаём подпапку "_image" в указанной папке
    folder_for_image = "_image"
    target_dir = os.path.join(folder_for_save, folder_for_image)
    if not os.path.isdir(target_dir):
        os.makedirs(target_dir)

    # добавляем имя файла (Revit требует базовое имя без расширения)

    base_filepath = os.path.join(target_dir, cur_doc.Title)

    # настраиваем экспорт
    ieopt = ImageExportOptions()
    ieopt.ZoomType = ZoomFitType.FitToPage
    ieopt.FitDirection = FitDirectionType.Horizontal
    ieopt.ImageResolution = ImageResolution.DPI_600
    ieopt.PixelSize = 512
    ieopt.ExportRange = ExportRange.CurrentView
    ieopt.HLRandWFViewsFileType = ImageFileType.JPEGLossless
    ieopt.ShadowViewsFileType = ImageFileType.JPEGLossless
    ieopt.FilePath = base_filepath

    # экспорт
    cur_doc.ExportImage(ieopt)

    result_path = base_filepath + ".jpg"
    # output.print_md(">> {}".format(result_path))
    return result_path
# --- main ---
# fam_docs = [d for d in revit.docs if d.IsFamilyDocument] 
# if not fam_docs:
#     script.exit()
# for fam_doc in fam_docs:
fam_doc = __revit__.ActiveUIDocument.Document
if not fam_doc.IsFamilyDocument:
    script.exit
path = fam_doc.PathName
uiapp.OpenAndActivateDocument(path)
folder_for_save = os.path.dirname(path)
# output.print_md(" - Работа с семейством {}".format(fam_doc.Title))
result_path = export_image(fam_doc, folder_for_save)
