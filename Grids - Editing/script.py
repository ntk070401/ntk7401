# -*- coding: utf-8 -*-
"""
pyRevit tool: Copy grid 2D extents, bubble visibility and elbows
from the active view to many other views.

Workflow:
1. Open the source view and adjust the selected grids as desired.
2. Preselect one or more grids in the source view.
3. Run this script.
4. Choose target views, then choose what to copy.

Revit rules respected by this tool:
    2D extent  ->  bubble  ->  elbow (leader)
A bubble must be visible before an elbow can be added.

2D and 3D extents:
The extent type is stored per grid end and per view. The tool reads the type of
each end in the source view and aligns the target views to it. A source that is
fully 3D keeps the targets 3D, because a 3D end is driven by the model extent.
A source with any 2D end switches the target ends to 2D before the curve is
written, so no change ever leaks into other views through a 3D end.

Why several passes are used:
Revit only reports the new grid end points after the document is regenerated.
An elbow position calculated from stale geometry lands in the wrong place,
which previously forced a second run. This version regenerates between passes
and then verifies and corrects every elbow inside the same run.

Compatibility: Revit 2022 - 2026, IronPython and CPython pyRevit engines.
"""

__title__ = "Edit grid "
__doc__ = "Apply 2D extents, bubble visibility and elbows"
__author__ = "KMNguyen"

import System
from pyrevit import revit, DB, forms


DOC = revit.doc
SOURCE_VIEW = revit.active_view

SHORT_CURVE_TOLERANCE = 0.0033   # feet, Revit minimum curve length
POSITION_TOLERANCE = 0.0026      # feet, about 1/32 inch
MAX_ELBOW_PASSES = 3

PLAN_VIEW_TYPES = set([
    DB.ViewType.FloorPlan,
    DB.ViewType.CeilingPlan,
    DB.ViewType.EngineeringPlan,
    DB.ViewType.AreaPlan,
])

SUPPORTED_VIEW_TYPES = PLAN_VIEW_TYPES.union(set([
    DB.ViewType.Elevation,
    DB.ViewType.Section,
    DB.ViewType.Detail,
]))

DATUM_ENDS = [DB.DatumEnds.End0, DB.DatumEnds.End1]


# ---------------------------------------------------------------- view filter

def are_parallel_views(view_a, view_b):
    try:
        direction_a = view_a.ViewDirection.Normalize()
        direction_b = view_b.ViewDirection.Normalize()
        dot = direction_a.DotProduct(direction_b)
        return abs(abs(dot) - 1.0) < 1e-6
    except Exception:
        return False


def is_dependent_view(view):
    """Dependent views inherit datum extents and cannot receive propagation."""
    try:
        return view.GetPrimaryViewId() != DB.ElementId.InvalidElementId
    except Exception:
        return False


def is_valid_target_view(view):
    if view.IsTemplate:
        return False
    if view.Id == SOURCE_VIEW.Id:
        return False
    if view.ViewType not in SUPPORTED_VIEW_TYPES:
        return False
    # Revit only propagates between views of the same type. A Floor Plan and a
    # Structural (Engineering) Plan are different types, which is what triggers
    # the "parallelViews" rejection.
    if view.ViewType != SOURCE_VIEW.ViewType:
        return False
    if is_dependent_view(view):
        return False
    try:
        if view.IsAssemblyView:
            return False
    except Exception:
        pass
    if not are_parallel_views(SOURCE_VIEW, view):
        return False
    return True


def get_compatible_views():
    views = []
    collector = DB.FilteredElementCollector(DOC).OfClass(DB.View)
    for view in collector:
        try:
            if is_valid_target_view(view):
                views.append(view)
        except Exception:
            continue
    views.sort(key=lambda x: x.Name)
    return views


class ViewItem(object):
    def __init__(self, view):
        self.view = view
        self.name = u"[{0}] {1}".format(view.ViewType, view.Name)


def get_selected_grids():
    grids = []
    for element in revit.get_selection().elements:
        if isinstance(element, DB.Grid):
            grids.append(element)
    return grids


# ------------------------------------------------------------------- geometry

def get_curves_in_view(grid, view):
    """Return grid curves in a view, preferring the view-specific extents."""
    for extent_type in [DB.DatumExtentType.ViewSpecific, DB.DatumExtentType.Model]:
        try:
            curves = list(grid.GetCurvesInView(extent_type, view))
        except Exception:
            curves = []
        if curves:
            return curves
    return []


def get_first_curve(grid, view):
    curves = get_curves_in_view(grid, view)
    return curves[0] if curves else None


def project_point_to_view_plane(point, plane_origin, plane_normal):
    """Project a point onto the datum plane used by the target view."""
    vector = point.Subtract(plane_origin)
    distance = vector.DotProduct(plane_normal)
    return point.Subtract(plane_normal.Multiply(distance))


def get_curve_direction(curve):
    try:
        start = curve.GetEndPoint(0)
        end = curve.GetEndPoint(1)
        return end.Subtract(start).Normalize()
    except Exception:
        return None


def get_end_index(datum_end, source_curve, target_curve):
    """Map a datum end to a curve end index, allowing for reversed curves.

    Some views return the grid curve in the opposite direction. Without this
    check the elbow of End0 can be applied to the physical opposite end.
    """
    index = 0 if datum_end == DB.DatumEnds.End0 else 1

    if source_curve is None or target_curve is None:
        return index

    source_direction = get_curve_direction(source_curve)
    target_direction = get_curve_direction(target_curve)
    if source_direction is None or target_direction is None:
        return index

    try:
        if source_direction.DotProduct(target_direction) < 0:
            return 1 - index
    except Exception:
        pass

    return index


def get_end_point(curve, index):
    if curve is None:
        return None
    try:
        return curve.GetEndPoint(index)
    except Exception:
        return None


# --------------------------------------------------------------- bubbles/2D

def is_bubble_visible(grid, datum_end, view):
    try:
        return grid.IsBubbleVisibleInView(datum_end, view)
    except Exception:
        return False


def set_bubble_visibility(grid, datum_end, view, visible):
    if visible:
        grid.ShowBubbleInView(datum_end, view)
    else:
        grid.HideBubbleInView(datum_end, view)


def get_extent_type(grid, datum_end, view):
    """Return Model (3D) or ViewSpecific (2D) for one grid end in a view."""
    try:
        return grid.GetDatumExtentTypeInView(datum_end, view)
    except Exception:
        return None


def set_extent_type(grid, datum_end, view, extent_type):
    try:
        grid.SetDatumExtentType(datum_end, view, extent_type)
        return True
    except Exception:
        return False


def force_view_specific(grid, view):
    """Switch both ends to 2D extents so edits stay local to this view."""
    for datum_end in DATUM_ENDS:
        set_extent_type(grid, datum_end, view, DB.DatumExtentType.ViewSpecific)


def match_extent_types(grid, view, source_extents):
    """Make the target view use the same 2D/3D end behaviour as the source.

    The extent type is stored per end, so End0 can be 3D while End1 is 2D.
    Returns True when a 2D curve should be written to the target view.
    """
    model = DB.DatumExtentType.Model
    view_specific = DB.DatumExtentType.ViewSpecific

    source_types = [source_extents.get("extent0"), source_extents.get("extent1")]
    known = [item for item in source_types if item is not None]

    # Source is fully 3D: keep the target 3D as well. A 3D end is driven by the
    # model extent, so writing a view-specific curve would be wrong here.
    if known and all(item == model for item in known):
        for datum_end in DATUM_ENDS:
            if get_extent_type(grid, datum_end, view) != model:
                set_extent_type(grid, datum_end, view, model)
        return False

    # Any 2D end means the target must be 2D before a curve can be written.
    # SetCurveInView always writes both ends, so both are switched to 2D.
    for datum_end in DATUM_ENDS:
        if get_extent_type(grid, datum_end, view) != view_specific:
            set_extent_type(grid, datum_end, view, view_specific)

    return True


def build_view_id_set(view):
    """ISet<ElementId> for PropagateToViews, safe on both pyRevit engines."""
    try:
        view_ids = System.Collections.Generic.HashSet[DB.ElementId]()
        view_ids.Add(view.Id)
        return view_ids
    except Exception:
        return None


def propagate_extents(grid, view):
    """Native Propagate Extents. Fails on views Revit considers invalid."""
    view_ids = build_view_id_set(view)
    if view_ids is None:
        raise Exception("View id set could not be created")
    grid.PropagateToViews(SOURCE_VIEW, view_ids)


def copy_extent_directly(grid, view, source_curve):
    """Fallback: rebuild the source line inside the target view's own plane.

    Projecting the endpoints onto the target view plane keeps the source length
    unchanged, avoids the "curve is not in the same plane" error, and never
    modifies the source view.
    """
    reference = get_first_curve(grid, view)
    if reference is None:
        raise Exception("Grid has no curve in this view")

    plane_origin = reference.GetEndPoint(0)

    try:
        plane_normal = view.ViewDirection.Normalize()
    except Exception:
        raise Exception("View direction is unavailable")

    start = project_point_to_view_plane(
        source_curve.GetEndPoint(0), plane_origin, plane_normal)
    end = project_point_to_view_plane(
        source_curve.GetEndPoint(1), plane_origin, plane_normal)

    if start.DistanceTo(end) < SHORT_CURVE_TOLERANCE:
        raise Exception("Projected extent is too short")

    force_view_specific(grid, view)
    grid.SetCurveInView(
        DB.DatumExtentType.ViewSpecific,
        view,
        DB.Line.CreateBound(start, end),
    )


def apply_extent(grid, view, source_curve):
    """Try the native propagation first, then the direct curve fallback."""
    try:
        propagate_extents(grid, view)
        return
    except Exception:
        pass

    if source_curve is None:
        raise Exception("Source view has no grid curve to copy")

    copy_extent_directly(grid, view, source_curve)


# ----------------------------------------------------------------- elbows

def get_leader(grid, datum_end, view):
    """Return the Leader object, or None when this end has no elbow.

    The API has no HasLeader method. GetLeader returns null when the end has no
    elbow, so probing with GetLeader is the reliable test.
    """
    try:
        return grid.GetLeader(datum_end, view)
    except Exception:
        return None


def get_source_leader(grid, datum_end, source_curve):
    """Capture the elbow and label position of one end in the source view.

    Positions are stored as offsets from the grid end point, so the elbow lands
    correctly even when a target view has different extents or sits on another
    level.
    """
    leader = get_leader(grid, datum_end, SOURCE_VIEW)
    if leader is None:
        return {"has_leader": False}

    try:
        elbow = leader.Elbow
        label_end = leader.End
    except Exception:
        return {"has_leader": False}

    data = {"has_leader": True, "elbow": elbow, "end": label_end}

    index = 0 if datum_end == DB.DatumEnds.End0 else 1
    anchor = get_end_point(source_curve, index)
    if anchor is not None and elbow is not None and label_end is not None:
        data["elbow_offset"] = elbow.Subtract(anchor)
        data["end_offset"] = label_end.Subtract(anchor)

    return data


def resolve_leader_points(grid, datum_end, view, leader_data, source_curve):
    """Rebuild the elbow and label points inside the target view plane."""
    reference = get_first_curve(grid, view)
    index = get_end_index(datum_end, source_curve, reference)
    anchor = get_end_point(reference, index)

    elbow_offset = leader_data.get("elbow_offset")
    end_offset = leader_data.get("end_offset")

    if anchor is not None and elbow_offset is not None and end_offset is not None:
        elbow = anchor.Add(elbow_offset)
        label_end = anchor.Add(end_offset)
    else:
        elbow = leader_data.get("elbow")
        label_end = leader_data.get("end")

    if reference is not None and elbow is not None and label_end is not None:
        try:
            plane_normal = view.ViewDirection.Normalize()
            plane_origin = reference.GetEndPoint(0)
            elbow = project_point_to_view_plane(elbow, plane_origin, plane_normal)
            label_end = project_point_to_view_plane(label_end, plane_origin, plane_normal)
        except Exception:
            pass

    return elbow, label_end


def ensure_leader(grid, datum_end, view, leader_data):
    """Create or remove the elbow so the target end matches the source end."""
    if leader_data.get("has_leader") is not True:
        if get_leader(grid, datum_end, view) is not None:
            try:
                grid.RemoveLeader(datum_end, view)
            except Exception:
                pass
        return False

    # An elbow can only exist where the bubble is displayed.
    if not is_bubble_visible(grid, datum_end, view):
        return False

    if get_leader(grid, datum_end, view) is None:
        # Equivalent to ticking "Add Elbow" on screen.
        grid.AddLeader(datum_end, view)

    return True


def place_leader(grid, datum_end, view, leader_data, source_curve):
    """Write the elbow and label position. Returns True when already correct."""
    leader = get_leader(grid, datum_end, view)
    if leader is None:
        raise Exception("Revit did not create an elbow in this view")

    elbow, label_end = resolve_leader_points(
        grid, datum_end, view, leader_data, source_curve)
    if elbow is None or label_end is None:
        raise Exception("Elbow position could not be calculated")

    try:
        current_elbow = leader.Elbow
        current_end = leader.End
        if (current_elbow is not None
                and current_end is not None
                and current_elbow.DistanceTo(elbow) < POSITION_TOLERANCE
                and current_end.DistanceTo(label_end) < POSITION_TOLERANCE):
            return True
    except Exception:
        pass

    # Writing to the live Leader object updates the model directly.
    leader.Elbow = elbow
    leader.End = label_end

    # Some API versions also expose an explicit write-back.
    if hasattr(grid, "SetLeader"):
        try:
            grid.SetLeader(datum_end, view, leader)
        except Exception:
            pass

    return False


def regenerate():
    """Refresh geometry so the next pass reads the updated grid end points."""
    try:
        DOC.Regenerate()
    except Exception:
        pass


# -------------------------------------------------------------------- UI

def choose_target_views():
    compatible = get_compatible_views()
    if not compatible:
        forms.alert(
            "No compatible target views were found.\n\n"
            "Target views must be the same view type as the active view, "
            "parallel to it, and not dependent views.",
            exitscript=True,
        )

    choice = forms.CommandSwitchWindow.show(
        ["Apply to all compatible views", "Select views manually"],
        message="Choose target views"
    )

    if choice == "Apply to all compatible views":
        return compatible
    if choice != "Select views manually":
        return []

    items = [ViewItem(view) for view in compatible]
    selected = forms.SelectFromList.show(
        items,
        title="Select target views",
        name_attr="name",
        multiselect=True,
        filterfunc=None,
    )
    if not selected:
        return []
    return [item.view for item in selected]


def choose_operation():
    options = [
        "Copy 2D extents + bubble visibility + elbow",
        "Copy 2D extents + bubble visibility",
        "Copy bubble visibility + elbow",
        "Copy bubble visibility only",
    ]
    return forms.CommandSwitchWindow.show(options, message="Choose operation")


# ------------------------------------------------------------------- main

def collect_source_data(grids):
    source_data = []
    for grid in grids:
        source_curve = get_first_curve(grid, SOURCE_VIEW)
        source_data.append({
            "grid": grid,
            "curve": source_curve,
            "extent0": get_extent_type(grid, DB.DatumEnds.End0, SOURCE_VIEW),
            "extent1": get_extent_type(grid, DB.DatumEnds.End1, SOURCE_VIEW),
            "end0": is_bubble_visible(grid, DB.DatumEnds.End0, SOURCE_VIEW),
            "end1": is_bubble_visible(grid, DB.DatumEnds.End1, SOURCE_VIEW),
            "leader0": get_source_leader(grid, DB.DatumEnds.End0, source_curve),
            "leader1": get_source_leader(grid, DB.DatumEnds.End1, source_curve),
        })
    return source_data


def visible_targets(grid, target_views):
    result = []
    for view in target_views:
        try:
            if not grid.CanBeVisibleInView(view):
                continue
        except Exception:
            pass
        result.append(view)
    return result


def main():
    if SOURCE_VIEW is None or SOURCE_VIEW.IsTemplate:
        forms.alert(
            "Run this tool from a normal plan, elevation, or section view.",
            exitscript=True,
        )

    if SOURCE_VIEW.ViewType not in SUPPORTED_VIEW_TYPES:
        forms.alert(
            "The active view type is not supported. Use a plan, elevation, or section view.",
            exitscript=True,
        )

    grids = get_selected_grids()
    if not grids:
        forms.alert(
            "Select one or more grids in the source view, then run the tool again.",
            exitscript=True,
        )

    target_views = choose_target_views()
    if not target_views:
        return

    operation = choose_operation()
    if not operation:
        return

    copy_extents = operation in [
        "Copy 2D extents + bubble visibility + elbow",
        "Copy 2D extents + bubble visibility",
    ]
    copy_leaders = operation in [
        "Copy 2D extents + bubble visibility + elbow",
        "Copy bubble visibility + elbow",
    ]

    failures = []

    with revit.Transaction("Copy grid extents, bubbles and elbows"):
        # The source grid is never modified.
        source_data = collect_source_data(grids)

        source_has_leader = False
        for data in source_data:
            if (data["leader0"].get("has_leader")
                    or data["leader1"].get("has_leader")):
                source_has_leader = True
                break

        jobs = []
        for data in source_data:
            for view in visible_targets(data["grid"], target_views):
                jobs.append((data, view))

        # Pass 1: 2D extents, then bubbles.
        for data, view in jobs:
            grid = data["grid"]

            if copy_extents:
                try:
                    # Align the 2D/3D end behaviour first. A target end left in
                    # 3D would otherwise change the grid in every view.
                    needs_curve = match_extent_types(grid, view, data)
                    if needs_curve:
                        apply_extent(grid, view, data["curve"])
                except Exception as error:
                    failures.append(u"Grid {0} / {1} / extent: {2}".format(
                        grid.Name, view.Name, error))

            try:
                set_bubble_visibility(grid, DB.DatumEnds.End0, view, data["end0"])
                set_bubble_visibility(grid, DB.DatumEnds.End1, view, data["end1"])
            except Exception as error:
                failures.append(u"Grid {0} / {1} / bubble: {2}".format(
                    grid.Name, view.Name, error))

        if copy_leaders:
            # Geometry must be refreshed before elbows are created, otherwise
            # AddLeader anchors on the old grid end point.
            regenerate()

            # Pass 2: create or remove elbows.
            elbow_jobs = []
            for data, view in jobs:
                grid = data["grid"]
                for datum_end in DATUM_ENDS:
                    key = "leader0" if datum_end == DB.DatumEnds.End0 else "leader1"
                    try:
                        if ensure_leader(grid, datum_end, view, data[key]):
                            elbow_jobs.append((data, view, datum_end, key))
                    except Exception as error:
                        failures.append(u"Grid {0} / {1} / elbow: {2}".format(
                            grid.Name, view.Name, error))

            # Pass 3: place the elbows, then verify and correct them.
            # Newly created elbows report their final anchor only after a
            # regeneration, which is why a verification loop replaces the
            # manual second run of the tool.
            pending = elbow_jobs
            for attempt in range(MAX_ELBOW_PASSES):
                if not pending:
                    break

                regenerate()
                still_pending = []

                for data, view, datum_end, key in pending:
                    grid = data["grid"]
                    try:
                        already_correct = place_leader(
                            grid, datum_end, view, data[key], data["curve"])
                        if not already_correct:
                            still_pending.append((data, view, datum_end, key))
                    except Exception as error:
                        if attempt == MAX_ELBOW_PASSES - 1:
                            failures.append(
                                u"Grid {0} / {1} / elbow: {2}".format(
                                    grid.Name, view.Name, error))

                pending = still_pending

            regenerate()

    # Keep successful runs quiet. Only warn when something really failed.
    if copy_leaders and not source_has_leader:
        failures.insert(
            0,
            "No elbow was found on the selected grids in the source view. "
            "Add the elbow in the source view first, then run the tool again.",
        )

    if failures:
        message = "Some grid operations could not be applied:\n\n"
        message += "\n".join(failures[:15])
        if len(failures) > 15:
            message += "\n...and {0} more.".format(len(failures) - 15)
        forms.alert(message, title="Grid tool completed with warnings")


if __name__ == "__main__":
    main()
