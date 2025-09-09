# -*- coding: utf-8 -*-
__title__ = "Из помещения\nв отделку"
__author__ = 'IliaNistratov'
__doc__ = """"""

from Autodesk.Revit.DB import (
    UV, Options, BuiltInCategory, FilteredElementCollector as FEC,
    BooleanOperationsType, BooleanOperationsUtils,
    GeometryElement, GeometryInstance, Solid, StorageType, Transaction, ViewDetailLevel)
from Autodesk.Revit.DB import LocationCurve, XYZ
from Autodesk.Revit.DB import Solid, GeometryCreationUtilities, XYZ,DirectShape, ElementId

from pyrevit import revit, output, script



uidoc = revit.uidoc
doc = __revit__.ActiveUIDocument.Document # type: ignore
user = __revit__.Application.Username # type: ignore
output = script.get_output()
output.close_others(all_open_outputs=True)
output.set_title(__title__)
output.set_width(900)
output.set_height(800)
lfy = output.linkify
vw = doc.ActiveView












def get_wall_endpoints(wall):
    """
    Возвращает начальную и конечную точку стены.
    :param wall: элемент стены (Wall)
    :return: (start_point, end_point) или (None, None) если не LocationCurve
    """
    loc = wall.Location
    if isinstance(loc, LocationCurve):
        curve = loc.Curve
        p1 = curve.GetEndPoint(0)
        p2 = curve.GetEndPoint(1)
        p3 = (p1 + p2) / 2.0  # середина линии
        if p1 and p2 and p3:
            print("{0} {1}:{3}:{2}".format(wall.Id, p1, p2, p3))
            return (p1, p2,p3)
    else:
        print("Стена ID {} не имеет LocationCurve".format(wall.Id))

        return (None, None, None)


def get_wall_inner_edge(wall):
    """
    Возвращает внутреннюю грань стены в виде пары точек (start, end).
    Работает с учетом ориентации стены и толщины.
    """
    loc = wall.Location
    if not isinstance(loc, LocationCurve):
        print("Стена ID {} не имеет LocationCurve".format(wall.Id))
        return (None, None)

    curve = loc.Curve
    p1 = curve.GetEndPoint(0)
    p2 = curve.GetEndPoint(1)

    # Направление осевой линии стены
    wall_direction = (p2 - p1).Normalize()

    # Перпендикуляр к направлению стены (в плоскости XY)
    perp = XYZ(-wall_direction.Y, wall_direction.X, 0)

    # Ширина (толщина) стены
    width = wall.Width

    # Смещение зависит от ориентации
    offset = perp * (width / 2.0)
    if wall.Flipped:
        offset = offset.Negate()

    p1 = p1 + offset
    p2 = p2 + offset
    p3 = (p1 + p2) / 2.0  # середина линии
    if p1 and p2 and p3:
        print("{0} {1}:{3}:{2}".format(wall.Id, p1, p2, p3))
    return (p1, p2,p3)

def show_solid(doc,solid, name_prefix="DebugSolid"):
    """
    Вставляет solid в модель через DirectShape в категории GenericModel
    """

    try:
        if solid or solid.Volume > 0:

            ds = DirectShape.CreateElement(doc, ElementId(BuiltInCategory.OST_GenericModel))
            ds.SetShape([solid])
            ds.Name = "{}".format(name_prefix)
            return ds
    except Exception as e:
        print("Ошибка при создании DirectShape: {}".format(e))
        return None


def build_solid_from_top_face(room, height=3000.0):
    """
    Строит Solid помещения, вытягивая верхнюю грань вниз на заданную высоту
    """


    top_face = None
    max_z = -99999

    # Находим верхнюю горизонтальную грань
    for face in get_solid_room(room)[0].Faces:
        normal = face.ComputeNormal(UV(0.5, 0.5))
        if abs(normal.Z - 1.0) < 0.01:  # почти вверх
            centroid = face.Evaluate(UV(0.5, 0.5))
            if centroid.Z > max_z:
                max_z = centroid.Z
                top_face = face

    if not top_face:
        print("Не удалось найти верхнюю грань помещения.")
        return None

    solids = []

    # Для каждого замкнутого контура на грани создаем Solid вытягиванием вниз
    for loop in top_face.GetEdgesAsCurveLoops():
        try:
            base_curve_loop = loop
            extrusion_dir = XYZ(0, 0, -1)
            solid = GeometryCreationUtilities.CreateExtrusionGeometry(
                [base_curve_loop],
                extrusion_dir,
                height / 304.8  # Перевод из мм в футы
            )
            if solid.Volume > 0:
                solids.append(solid)
        except Exception as e:
            print("Ошибка при создании Solid: {}".format(e))

    # Можно объединить Solids, если их несколько
    if not solids:
        return None

    final = solids[0]
    for s in solids[1:]:
        final = BooleanOperationsUtils.ExecuteBooleanOperation(final, s, BooleanOperationsType.Union)

    return final



def get_rooms(d):
    return FEC(d, vw.Id).OfCategory(BuiltInCategory.OST_Rooms).ToElements()


def get_walls(d):
    return FEC(d, vw.Id).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType().ToElements()



def intersect_solid(rs, ws):
    try:
        inter = BooleanOperationsUtils.ExecuteBooleanOperation(rs, ws, BooleanOperationsType.Intersect)
        if inter and inter.Volume > 0.00000000000000000001:
            
            return inter.Volume
        else:
            return False
    except Exception as e:
        print("Ошибка при пересечении: {}".format(e))
        return False

def get_solids(el):
    lst = []
    result = []
    op = Options()
    op.DetailLevel = ViewDetailLevel.Fine
    ge = el.get_Geometry(op)
    if isinstance(ge, GeometryElement):
        lst.append(ge)
    while len(lst) > 0:
        current = lst.pop(0)
        for go in current.GetEnumerator():
            if isinstance(go, GeometryInstance):
                lst.append(go.GetInstanceGeometry())
            elif isinstance(go, Solid) and go.Volume > 0:
                result.append(go)
    return result


def get_solid_room(room):
    lst = []
    result = []
    ge = room.ClosedShell
    if not ge:
        print("У помещения ID {} нет замкнутой оболочки".format(room.Id))
        return result
    if isinstance(ge, GeometryElement):
        lst.append(ge)
    while len(lst) > 0:
        current = lst.pop(0)
        for go in current.GetEnumerator():
            if isinstance(go, GeometryInstance):
                lst.append(go.GetInstanceGeometry())
            elif isinstance(go, Solid) and go.Volume > 0:
                result.append(go)
    return result

walls = get_walls(doc)
rooms = get_rooms(doc)



with Transaction(doc, "Запись имени и номера помещения в стены") as t:
    t.Start()
    intersected_dict = {}

    for r in rooms:
        print("Помещение {}".format(lfy(r.Id)))
        solid_rooms = get_solid_room(r)
        if not solid_rooms:
            continue

        room_name = r.LookupParameter("Имя")
        room_number = r.LookupParameter("Номер")
        room_name_val = room_name.AsString() if room_name else ""
        room_number_val = room_number.AsString() if room_number else ""

        for w in walls:  
            solid_walls = get_solids(w)
            if not solid_walls:
                continue

            found = False
            for sr in solid_rooms:
                for sw in solid_walls:
                    vol = intersect_solid(sr, sw)
                    if vol:
                        print("  >>> Найдено пересечение между помещением {} и стеной {} (объем: {:.4f})".format(
                            lfy(r.Id), lfy(w.Id), vol))

                        if r.Id not in intersected_dict:
                            intersected_dict[r.Id] = []
                        intersected_dict[r.Id].append(w)

                        # Запись параметров
                        p_name = w.LookupParameter("Помещения_Имя")
                        p_number = w.LookupParameter("Помещения_Номер")

                        if p_name and p_name.StorageType == StorageType.String:
                            p_name.Set(room_name_val)
                        if p_number and p_number.StorageType == StorageType.String:
                            p_number.Set(room_number_val)

                        found = True
                        break
                if found:
                    break

    t.Commit()