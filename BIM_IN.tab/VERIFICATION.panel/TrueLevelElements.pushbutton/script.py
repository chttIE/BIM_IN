# -*- coding: utf-8 -*-

import os
from Autodesk.Revit.DB import (
    FilteredElementCollector,
    BuiltInParameter,
    Level,
    FamilyInstance,
    ElementId,
    TemporaryViewMode,
    OverrideGraphicSettings,
    Color,
    FillPatternElement
)
from pyrevit import revit, forms
from System.Collections.Generic import List


doc = revit.doc
uidoc = revit.uidoc


level_params = [
    BuiltInParameter.LEVEL_PARAM,
    BuiltInParameter.FAMILY_LEVEL_PARAM,
    BuiltInParameter.INSTANCE_REFERENCE_LEVEL_PARAM,
]


def to_element_id_list(ids):
    result = List[ElementId]()
    for eid in ids:
        if eid and eid != ElementId.InvalidElementId:
            result.Add(eid)
    return result


def clean_name(name):
    return name.strip().lower() if name else ""


def get_element_level(el):
    for bip in level_params:
        try:
            p = el.get_Parameter(bip)
            if p and p.StorageType == 2:
                level = doc.GetElement(p.AsElementId())
                if isinstance(level, Level):
                    return level
        except:
            pass

    try:
        level = doc.GetElement(el.LevelId)
        if isinstance(level, Level):
            return level
    except:
        pass

    try:
        if el.Host:
            return get_element_level(el.Host)
    except:
        pass

    return None


def collect_view_elements():
    result = []
    collector = (
        FilteredElementCollector(doc, doc.ActiveView.Id)
        .WhereElementIsNotElementType()
    )

    for el in collector:
        try:
            if not el.Category:
                continue
            if isinstance(el, FamilyInstance) and el.SuperComponent:
                continue
            result.append(el)
        except:
            pass

    return result


def reset_temp(view):
    if view.IsTemporaryHideIsolateActive():
        view.DisableTemporaryViewMode(TemporaryViewMode.TemporaryHideIsolate)


# -----------------------------
# ЛОГИКА
# -----------------------------

def isolate_by_level(level_name):
    view = doc.ActiveView
    target = clean_name(level_name)

    ids = []

    for el in collect_view_elements():
        lvl = get_element_level(el)
        if lvl and clean_name(lvl.Name) == target:
            ids.append(el.Id)

    ids = to_element_id_list(ids)

    with revit.Transaction("Изоляция"):
        reset_temp(view)
        if ids.Count > 0:
            view.IsolateElementsTemporary(ids)

    uidoc.RefreshActiveView()


def isolate_no_level():
    view = doc.ActiveView

    ids = []

    for el in collect_view_elements():
        if not get_element_level(el):
            ids.append(el.Id)

    ids = to_element_id_list(ids)

    with revit.Transaction("Без уровня"):
        reset_temp(view)
        if ids.Count > 0:
            view.IsolateElementsTemporary(ids)

    uidoc.RefreshActiveView()


def select_by_level(level_name):
    target = clean_name(level_name)

    current_ids = list(uidoc.Selection.GetElementIds())

    new_ids = []

    for el in collect_view_elements():
        lvl = get_element_level(el)
        if lvl and clean_name(lvl.Name) == target:
            new_ids.append(el.Id)

    new_ids_set = set([i.IntegerValue for i in new_ids])
    current_ids_set = set([i.IntegerValue for i in current_ids])

    # 🔥 toggle
    if new_ids_set == current_ids_set:
        uidoc.Selection.SetElementIds(List[ElementId]())  # сброс
    else:
        uidoc.Selection.SetElementIds(to_element_id_list(new_ids))

def select_no_level():
    current_ids = list(uidoc.Selection.GetElementIds())

    new_ids = []

    for el in collect_view_elements():
        if not get_element_level(el):
            new_ids.append(el.Id)

    new_ids_set = set([i.IntegerValue for i in new_ids])
    current_ids_set = set([i.IntegerValue for i in current_ids])

    if new_ids_set == current_ids_set:
        uidoc.Selection.SetElementIds(List[ElementId]())
    else:
        uidoc.Selection.SetElementIds(to_element_id_list(new_ids))


def color_by_level(level_name):
    view = doc.ActiveView
    target = clean_name(level_name)

    ogs = OverrideGraphicSettings()
    ogs.SetProjectionLineColor(Color(255, 0, 0))

    # solid fill
    solid = None
    for fp in FilteredElementCollector(doc).OfClass(FillPatternElement):
        if fp.GetFillPattern().IsSolidFill:
            solid = fp.Id
            break

    if solid:
        ogs.SetSurfaceForegroundPatternId(solid)
        ogs.SetSurfaceForegroundPatternColor(Color(255, 0, 0))

    with revit.Transaction("Окрасить"):
        for el in collect_view_elements():
            lvl = get_element_level(el)
            if lvl and clean_name(lvl.Name) == target:
                view.SetElementOverrides(el.Id, ogs)

def reset_colors():
    view = doc.ActiveView

    with revit.Transaction("Сброс окраски"):
        for el in collect_view_elements():
            view.SetElementOverrides(el.Id, OverrideGraphicSettings())

    uidoc.RefreshActiveView()


def reset_view():
    view = doc.ActiveView
    with revit.Transaction("Сброс"):
        reset_temp(view)
    uidoc.RefreshActiveView()


# -----------------------------
# UI
# -----------------------------

class Window(forms.WPFWindow):
    def __init__(self, xaml, levels):
        forms.WPFWindow.__init__(self, xaml)

        for lvl in levels:
            self.level_combo.Items.Add(lvl)

        self.level_combo.SelectedIndex = 0

        self.show_level_btn.Click += self.on_show
        self.no_level_btn.Click += self.on_no_level
        self.reset_btn.Click += self.on_reset
        self.select_level_btn.Click += self.on_select
        self.select_no_level_btn.Click += self.on_select_no
        self.color_level_btn.Click += self.on_color
        self.reset_color_btn.Click += self.on_reset_color

    def get_lvl(self):
        return self.level_combo.SelectedItem

    def on_show(self, s, a):
        lvl = self.get_lvl()
        isolate_by_level(lvl.Name)

    def on_no_level(self, s, a):
        isolate_no_level()

    def on_select(self, s, a):
        lvl = self.get_lvl()
        select_by_level(lvl.Name)

    def on_select_no(self, s, a):
        select_no_level()

    def on_color(self, s, a):
        lvl = self.get_lvl()
        color_by_level(lvl.Name)

    def on_reset(self, s, a):
        reset_view()

    def on_reset_color(self, s, a):
        reset_colors()
# -----------------------------
# RUN
# -----------------------------

levels = sorted(
    list(FilteredElementCollector(doc).OfClass(Level)),
    key=lambda x: x.Elevation
)

xaml = os.path.join(os.path.dirname(__file__), "window.xaml")

Window(xaml, levels).ShowDialog()