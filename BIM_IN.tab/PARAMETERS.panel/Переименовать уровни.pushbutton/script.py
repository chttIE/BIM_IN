# -*- coding: utf-8 -*-

import os

import clr
clr.AddReference("PresentationFramework")
clr.AddReference("PresentationCore")
clr.AddReference("WindowsBase")

from Autodesk.Revit.DB import FilteredElementCollector, Level, Transaction
from pyrevit import forms

try:
    from Autodesk.Revit.DB import UnitTypeId, UnitUtils
    HAS_UNIT_TYPE_ID = True
except:
    from Autodesk.Revit.DB import DisplayUnitType, UnitUtils
    HAS_UNIT_TYPE_ID = False

from System.Collections.ObjectModel import ObservableCollection
from System.Windows.Input import Key, ModifierKeys, Keyboard


doc = __revit__.ActiveUIDocument.Document

try:
    unicode
except NameError:
    unicode = str

TEMP_PREFIX = u"__BIMIN_TEMP_LEVEL_RENAME__"


class LevelRow(object):
    def __init__(self, level):
        self.Level = level
        self.Id = level.Id.IntegerValue
        self.Elevation = level.Elevation
        self.ElevationText = format_elevation(level.Elevation)
        self.Name = level.Name
        self.NewName = level.Name

    @property
    def Changed(self):
        return self.Name != self.NewName


def format_elevation(elevation_feet):
    try:
        if HAS_UNIT_TYPE_ID:
            mm = UnitUtils.ConvertFromInternalUnits(elevation_feet, UnitTypeId.Millimeters)
        else:
            mm = UnitUtils.ConvertFromInternalUnits(elevation_feet, DisplayUnitType.DUT_MILLIMETERS)
    except:
        mm = elevation_feet * 304.8

    mm_int = int(round(mm))
    sign = u"-" if mm_int < 0 else u""
    value = abs(mm_int)
    text = u"{:,}".format(value).replace(",", " ")
    return sign + text


def clean_name(value):
    if value is None:
        return u""
    return unicode(value).strip()


def validate_rows(rows):
    errors = []
    new_names = {}
    changed_ids = set()

    for row in rows:
        row.NewName = clean_name(row.NewName)
        if row.Changed:
            changed_ids.add(row.Id)

    for row in rows:
        name = row.NewName
        if not name:
            errors.append(u"Уровень '{}' получит пустое имя.".format(row.Name))
            continue

        lowered = name.lower()
        if lowered in new_names:
            errors.append(
                u"Повтор имени: '{}' задан для '{}' и '{}'.".format(
                    name,
                    new_names[lowered].Name,
                    row.Name
                )
            )
        else:
            new_names[lowered] = row

    for level in FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType():
        level_id = level.Id.IntegerValue
        if level_id in changed_ids:
            continue

        if level.Name.lower() in new_names:
            target_row = new_names[level.Name.lower()]
            if target_row.Id != level_id:
                errors.append(
                    u"Имя '{}' уже занято уровнем Id {}.".format(
                        target_row.NewName,
                        level_id
                    )
                )

    return errors


def get_changed_rows(rows):
    return [row for row in rows if row.Changed]


def apply_renames(rows):
    changed = get_changed_rows(rows)
    if not changed:
        forms.alert(u"Нет изменений для применения.", title=u"Переименование уровней")
        return False

    errors = validate_rows(rows)
    if errors:
        forms.alert(
            u"\n".join(errors[:12]),
            title=u"Проверьте имена уровней",
            warn_icon=True
        )
        return False

    tx = Transaction(doc, u"Переименование уровней")
    tx.Start()
    try:
        for row in changed:
            row.Level.Name = u"{}_{}".format(TEMP_PREFIX, row.Id)

        for row in changed:
            row.Level.Name = row.NewName

        tx.Commit()
    except Exception as exc:
        if tx.HasStarted():
            tx.RollBack()
        forms.alert(
            u"Не удалось переименовать уровни:\n{}".format(exc),
            title=u"Ошибка",
            warn_icon=True
        )
        return False

    forms.alert(
        u"Переименовано уровней: {}".format(len(changed)),
        title=u"Готово"
    )
    return True


class LevelRenameWindow(forms.WPFWindow):
    def __init__(self, xaml_file, rows):
        forms.WPFWindow.__init__(self, xaml_file)
        self.rows = rows
        self.undo_stack = []
        self.levels_grid.ItemsSource = rows
        self.replace_btn.Click += self.on_replace
        self.apply_btn.Click += self.on_apply
        self.close_btn.Click += self.on_close
        self.PreviewKeyDown += self.on_preview_key_down

    def push_undo(self, snapshot):
        if not snapshot:
            return

        self.undo_stack.append(snapshot)
        if len(self.undo_stack) > 20:
            self.undo_stack.pop(0)

    def undo_last_action(self):
        if not self.undo_stack:
            return False

        snapshot = self.undo_stack.pop()
        for row, old_name in snapshot:
            row.NewName = old_name

        self.levels_grid.Items.Refresh()
        return True

    def on_preview_key_down(self, sender, args):
        if args.Key != Key.Z or Keyboard.Modifiers != ModifierKeys.Control:
            return

        # Let active text boxes use their own undo while the user is typing.
        try:
            focused = Keyboard.FocusedElement
            if focused and focused.CanUndo:
                return
        except:
            pass

        self.levels_grid.CommitEdit()
        self.levels_grid.CommitEdit()
        if self.undo_last_action():
            args.Handled = True

    def get_replace_rows(self):
        selected_rows = []

        try:
            for cell_info in self.levels_grid.SelectedCells:
                row = cell_info.Item
                if row and row not in selected_rows:
                    selected_rows.append(row)
        except:
            pass

        if not selected_rows:
            try:
                for row in self.levels_grid.SelectedItems:
                    if row and row not in selected_rows:
                        selected_rows.append(row)
            except:
                pass

        if selected_rows:
            return selected_rows

        return list(self.rows)

    def on_replace(self, sender, args):
        self.levels_grid.CommitEdit()
        self.levels_grid.CommitEdit()

        find_text = unicode(self.find_box.Text)
        replace_text = unicode(self.replace_box.Text)

        if find_text == u"":
            forms.alert(
                u"Заполните поле 'Искать'.",
                title=u"Автозамена",
                warn_icon=True
            )
            return

        changed_count = 0
        snapshot = []
        for row in self.get_replace_rows():
            current_name = unicode(row.NewName)
            if find_text in current_name:
                snapshot.append((row, current_name))
                row.NewName = current_name.replace(find_text, replace_text)
                changed_count += 1

        self.push_undo(snapshot)
        self.levels_grid.Items.Refresh()

    def on_apply(self, sender, args):
        self.levels_grid.CommitEdit()
        self.levels_grid.CommitEdit()
        if apply_renames(list(self.rows)):
            self.Close()

    def on_close(self, sender, args):
        self.Close()


levels = sorted(
    list(FilteredElementCollector(doc).OfClass(Level).WhereElementIsNotElementType()),
    key=lambda x: x.Elevation
)

if not levels:
    forms.alert(u"В проекте нет уровней.", title=u"Переименование уровней")
else:
    rows = ObservableCollection[object]()
    for level in levels:
        rows.Add(LevelRow(level))

    xaml = os.path.join(os.path.dirname(__file__), "window.xaml")
    LevelRenameWindow(xaml, rows).ShowDialog()
