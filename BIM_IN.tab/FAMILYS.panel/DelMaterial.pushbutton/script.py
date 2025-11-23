# -*- coding: utf-8 -*-
__title__ = 'Открыть семейства'
__doc__ = ""
__highlight__ = 'new'
__context__ = "zero-doc"

import os
import re
import clr
clr.AddReference('System')

from System.Collections.Generic import List  # type: ignore
from Autodesk.Revit.DB import (
    BuiltInCategory, Category, ConnectorElement, Document, ElementId, ExportRange, FailureProcessingResult, FailureSeverity, FilteredElementCollector, FitDirectionType, IFailuresPreprocessor, IFailuresProcessor, ImageExportOptions, ImageFileType, ImageResolution, Material, Transaction,
    XYZ, DisplayStyle, ElementTypeGroup, View, View3D, ViewDetailLevel,
    ViewDuplicateOption, SaveAsOptions, ModelPathUtils, ZoomFitType
)
from pyrevit import forms, revit, script

# удобный алиас
FEC = FilteredElementCollector

uidoc = __revit__.ActiveUIDocument  # type: ignore
app = __revit__.Application          # type: ignore
active_doc = uidoc.Document          # type: ignore

output = script.get_output()
script.get_output().close_others(all_open_outputs=True)


# --- utils ---

def Hidden_connectors(cur_doc,v):
    output.print_md(" >> Прячем коннекторы")
    connector = FEC(cur_doc,v.Id).OfCategory(BuiltInCategory.OST_ConnectorElem).ToElementIds()
    if connector: 

        v.HideElementsTemporary(connector)
        output.print_md(" >> Спрятал")

    return


def setting_view(cur_doc, v):
    """Настраиваю вид (нужна активная транзакция снаружи!)."""
    output.print_md("- Задаю настройки вида")
    v.AreAnnotationCategoriesHidden = True
    v.AreImportCategoriesHidden = True
    v.ArePointCloudsHidden = True
    v.CropBoxActive = False
    v.CropBoxVisible = False
    v.DetailLevel = ViewDetailLevel.Fine
    v.DisplayStyle = DisplayStyle.Shading
    v.IsSectionBoxActive = False
    try:
        v.ToggleToIsometric()
        v.OrientTo(XYZ(2.27101522431622, 2.32807317165811, -2.34309571423024))
    except:
        output.print_md("- Не смог включить изометрию")
    v.RevealConstraintsMode = False
    Hidden_connectors(cur_doc, v)


def get_preview_view(cur_doc):
    """Возвращает 3D вид 'Миниатюра' либо False."""
    return next((view for view in FEC(cur_doc).OfClass(View3D) if view.Name == 'Миниатюра'), False)


def make_view_for_save(cur_doc, cur_uidoc):
    """
    Создаёт/находит вид 'Миниатюра', делает его активным,
    затем настраивает и скрывает коннекторы. Возвращает True/False.
    """
    # T1: создать/найти вид
    with Transaction(cur_doc, 'pyIN | Создал/нашёл вид') as t1:
        t1.Start()
        v = get_preview_view(cur_doc)
        if not v:
            output.print_md("- Вид не найден. Создаю")
            view_type_3d_id = cur_doc.GetDefaultElementTypeId(ElementTypeGroup.ViewType3D)
            v = View3D.CreateIsometric(cur_doc, view_type_3d_id)
            v.Name = "Миниатюра"
            output.print_md("- Вид создан")
        else:
            output.print_md("- Вид найден.")
        t1.Commit()

    # Активируем только что созданный/найденный вид
    try:
        cur_uidoc.RequestViewChange(v)
        output.print_md("- Сделал активным вид '{}'".format(v.Name))
    except Exception as e:
        output.print_md("! Не удалось сделать вид активным (RequestViewChange) {}".format(e))
    with Transaction(cur_doc, 'pyIN | задаем настройки вида') as t2:
        t2.Start()
        setting_view(cur_doc, v)  # внутри вызовет hidden_connectors
        t2.Commit()


    return True

import tempfile
import os

def save_fam(cur_doc, folder, preffix=None, use_preview=True, compact=True):
    output.print_md("- Сохранение семейства")
    name = "{}{}.rfa".format(preffix or "", cur_doc.Title)
    target_path = os.path.join(folder, name)
    output.print_md(">> Полный путь до семейства: {}".format(target_path))

    cur_path = getattr(cur_doc, "PathName", "") or ""
    try:
        same_target = os.path.normcase(os.path.abspath(cur_path)) == os.path.normcase(os.path.abspath(target_path))
    except:
        same_target = False

    from Autodesk.Revit.DB import SaveAsOptions, ModelPathUtils
    opts = SaveAsOptions()
    opts.Compact = bool(compact)
    opts.MaximumBackups = 1
    opts.OverwriteExistingFile = True
    if use_preview:

        view_for_save = get_preview_view(cur_doc)
        if view_for_save:
            output.print_md(">> Используем вид {}".format(view_for_save.Name))
            opts.PreviewViewId = view_for_save.Id

    if same_target:
        # 1) Сохраняем во временный файл и сообщаем вызывающему коду, что нужна отложенная замена
        tmp_dir = tempfile.mkdtemp(prefix="rfa_tmp_")
        tmp_path = os.path.join(tmp_dir, os.path.basename(target_path))
        mp_tmp = ModelPathUtils.ConvertUserVisiblePathToModelPath(tmp_path)
        try:
            # output.print_md("- Целевой путь совпадает. Сохраняю во временный: {}".format(tmp_path))
            cur_doc.SaveAs(mp_tmp, opts)
            # output.print_md("- Временное сохранение успех. Нужна отложенная замена после закрытия документа.")
            return True, {"deferred_replace": (tmp_dir, tmp_path, target_path)}
        except Exception as e:
            output.print_md(">> ! Ошибка при временном сохранении: {}".format(e))
            # Постараемся прибрать временное
            try:
                if os.path.exists(tmp_path):
                    os.remove(tmp_path)
                if os.path.isdir(tmp_dir):
                    os.rmdir(tmp_dir)
            except:
                pass
            return False, None
    else:
        # 2) Прямое сохранение в новый путь
        mp_target = ModelPathUtils.ConvertUserVisiblePathToModelPath(target_path)
        try:
            # output.print_md("- Сохраняю напрямую (без compact) ...")
            cur_doc.SaveAs(mp_target, opts)
            output.print_md(">> Сохранено.")
            return True, None
        except Exception as e:
            output.print_md(">> ! Ошибка при сохранении: {}".format(e))
            return False, None

def activate_any_project_doc():
    """Ищет открытый не-family документ, активирует его и переключает на безопасный вид."""
    for d in app.Documents:
        if d.IsFamilyDocument:
            continue
        path = (getattr(d, "PathName", "") or "").strip()
        if not path:
            continue  # без пути OpenAndActivateDocument активировать нельзя
        try:
            host_uidoc = __revit__.OpenAndActivateDocument(path)  # активирует уже открытый проект
            # # выбираем вид: сначала любой 3D, иначе первый не-шаблонный
            # v = next((v for v in FEC(d).OfClass(View3D)), None) \
            #     or next((v for v in FEC(d).OfClass(View) if not v.IsTemplate), None)
            # if v:
            #     host_uidoc.RequestViewChange(v)  # <-- корректная смена активного вида
            return True
        except:
            print("Jib,rf")
            continue
    return False

from System.IO import File, Directory
from System.Threading import Thread
import stat
import os


class WarningSwallower_BC(IFailuresPreprocessor):
    def PreprocessFailures(self, failuresAccessor):
        try:
            fail_messages = failuresAccessor.GetFailureMessages()
            for fail_m in fail_messages:
                try:
                    if fail_m.GetSeverity() != FailureSeverity.Warning:
                        if fail_m.HasResolutions():
                            failuresAccessor.ResolveFailure(fail_m)
                    else:
                        failuresAccessor.DeleteWarning(fail_m)
                except:
                    pass
            return FailureProcessingResult.Continue
        except:
            return FailureProcessingResult.ProceedWithRollBack

class GlobalFailureSwallower(IFailuresProcessor):
    def ProcessFailures(self, failuresAccessor):
        try:
            fail_messages = failuresAccessor.GetFailureMessages()
            for fail_m in fail_messages:
                try:
                    msg_text = fail_m.GetDescriptionText()
                    print("  - Обнаружена ошибка: **{}**".format(msg_text))

                    if fail_m.GetSeverity() != FailureSeverity.Warning:
                        if fail_m.HasResolutions():
                            print("    → Автоматически применено решение.")
                            failuresAccessor.ResolveFailure(fail_m)
                    else:
                        print("    → Предупреждение удалено.")
                        failuresAccessor.DeleteWarning(fail_m)
                except Exception as inner_e:
                    print("    ⚠️ Не удалось обработать: {}".format(inner_e))
            return FailureProcessingResult.ProceedWithCommit
        except Exception as e:
            print("⛔ Ошибка при обработке: {}".format(e))
            return FailureProcessingResult.ProceedWithRollBack

    def Dismiss(self, document):
        pass

class OFFGlobalFailureSwallower(IFailuresProcessor):
    def ProcessFailures(self, failuresAccessor):
        pass

    def Dismiss(self, document):
        pass



def _retry(action, attempts=5, delay_ms=200):
    for i in range(attempts):
        try:
            return action()
        except Exception as e:
            if i == attempts - 1:
                raise
            Thread.Sleep(delay_ms)

def finalize_replace(tmp_dir, tmp_path, target_path):
    """
    Безопасная замена: удаляем целевой файл (с ретраями), затем перемещаем tmp -> target.
    Если Move проваливается, пробуем Copy. Если и это не вышло — оставляем tmp и сообщаем путь.
    """
    # 1) снимаем read-only у цели (если есть)
    if os.path.exists(target_path):
        try:
            os.chmod(target_path, stat.S_IWRITE)
        except:
            pass

    # 2) удаляем цель с ретраями (иногда антивирус/индексатор держит дескриптор)
    if os.path.exists(target_path):
        try:
            _retry(lambda: File.Delete(target_path))
        except Exception as e:
            output.print_md(">> ! Не удалось удалить целевой файл: {}. Пробую продолжить.".format(e))

    # 3) перемещаем tmp -> target с ретраями
    try:
        _retry(lambda: File.Move(tmp_path, target_path))
        output.print_md(">> Файл заменён перемещением.")
        moved = True
    except Exception as move_err:
        output.print_md(">> ! Move не удался: {}. Пробую Copy.".format(move_err))
        moved = False
        try:
            _retry(lambda: File.Copy(tmp_path, target_path, True))
            # если скопировали — можно удалить tmp
            try:
                File.Delete(tmp_path)
            except:
                pass
            output.print_md(">> Файл заменён копированием.")
            moved = True
        except Exception as copy_err:
            output.print_md(">> !! Copy тоже не удался: {}".format(copy_err))
            # ВАЖНО: оставляем tmp как резервную копию и не чистим tmp_dir
            output.print_md(">> !! РЕЗЕРВ СОХРАНЁН: {}".format(tmp_path))
            return False

    # 4) уборка временной папки (только если замена успешна)
    try:
        if os.path.exists(tmp_path):
            try:
                File.Delete(tmp_path)
            except:
                pass
        if os.path.isdir(tmp_dir):
            # удаляем каталог, если пустой
            try:
                Directory.Delete(tmp_dir, True)
            except:
                pass
    except:
        pass

    output.print_md(">> Оригинал заменён: {}".format(target_path))
    output.print_md(">> Сохранение завершенно")
    return True

def _sanitize(name):
    """Убирает недопустимые для имени файла/папки символы."""
    # Windows-запрещённые символы: \ / : * ? " < > | и управляющие
    name = re.sub(r'[\\/:*?"<>|]+', '_', name)
    name = re.sub(r'[\x00-\x1f]+', '_', name)
    name = name.strip().rstrip('. ')  # нельзя заканчивать точкой/пробелом
    return name or "untitled"


def del_material(cur_doc):
    ignore_materials_name = ["ЗО", "Зона обслуживания"]
    output.print_md("- Удаление материалов")
    with Transaction(cur_doc, "RemoveMaterial") as t:
        t.Start()
        for mat in FilteredElementCollector(cur_doc).OfClass(Material).ToElements():
            # исправленная проверка — сравнение без учёта регистра
            if mat.Name.upper() in [name.upper() for name in ignore_materials_name]:
                output.print_md(">> Пропустил {}".format(mat.Name))
                continue
            try:
                output.print_md(">> Удалил {}".format(mat.Name))
                cur_doc.Delete(mat.Id)
            except:
                output.print_md(">> НЕ смог удалить материал: {}".format(mat.Name))
                continue
        t.Commit()

def del_material(cur_doc):
    ignore_materials_name = ["ЗО", "Зона обслуживания"]
    output.print_md("- Удаление материалов")
    with Transaction(cur_doc, "RemoveMaterial") as t:
        t.Start()
        for mat in FilteredElementCollector(cur_doc).OfClass(Material).ToElements():
            # исправленная проверка — сравнение без учёта регистра
            if mat.Name.upper() in [name.upper() for name in ignore_materials_name]:
                output.print_md(">> Пропустил {}".format(mat.Name))
                continue
            try:
                output.print_md(">> Удалил {}".format(mat.Name))
                cur_doc.Delete(mat.Id)
            except:
                output.print_md(">> НЕ смог удалить материал: {}".format(mat.Name))
                continue
        t.Commit()

# --- main ---
#MAIN

fam_docs = [d for d in revit.docs if d.IsFamilyDocument] 
if not fam_docs:
    script.exit()

for fam_doc in fam_docs:
    output.print_md(" - Работа с семейством {}".format(fam_doc.Title))
    del_material(fam_doc)