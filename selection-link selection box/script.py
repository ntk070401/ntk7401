from pyrevit import revit, DB, forms
from Autodesk.Revit.DB import (
    XYZ, BoundingBoxXYZ, Transaction, View3D, ViewFamily, ViewFamilyType,
    FilteredElementCollector
)
from Autodesk.Revit.UI.Selection import ObjectType

doc = revit.doc
uidoc = revit.uidoc

# Optional padding around the combined bounding box, in feet
PADDING_FT = 1.0

# Revit's default 3D view naming convention is "{3D - <username>}",
# one per logged-in account. Reuse that same view instead of a fixed name.
username = doc.Application.Username
VIEW_NAME = "{{3D - {}}}".format(username)


# 1. Get references to linked elements: prefer current pre-selection,
#    fall back to interactive pick if nothing is selected.
refs = list(uidoc.Selection.GetReferences())

if not refs:
    try:
        refs = list(uidoc.Selection.PickObjects(
            ObjectType.LinkedElement,
            "Select elements inside the linked model, then press Finish"
        ))
    except Exception:
        forms.alert("Selection cancelled.", exitscript=True)

# Keep only references that actually point into a linked model
refs = [r for r in refs if r.LinkedElementId is not None and r.LinkedElementId != DB.ElementId.InvalidElementId]

if not refs:
    forms.alert(
        "No linked elements selected. Ctrl+Click elements inside a "
        "linked model, or select them when prompted.",
        exitscript=True
    )

# 2. Compute the combined bounding box in host (project) coordinates
overall_min = None
overall_max = None

for ref in refs:
    link_instance = doc.GetElement(ref.ElementId)
    if not isinstance(link_instance, DB.RevitLinkInstance):
        continue

    link_doc = link_instance.GetLinkDocument()
    if link_doc is None:
        continue

    linked_elem = link_doc.GetElement(ref.LinkedElementId)
    if linked_elem is None:
        continue

    bbox = linked_elem.get_BoundingBox(None)
    if bbox is None:
        continue

    # GetTotalTransform accounts for the link's own position/rotation
    # as well as any nested link transforms
    transform = link_instance.GetTotalTransform()

    corners = [
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Min.Z),
        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
        XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
        XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z),
        XYZ(bbox.Max.X, bbox.Max.Y, bbox.Max.Z),
    ]

    for corner in corners:
        pt = transform.OfPoint(corner)
        if overall_min is None:
            overall_min = pt
            overall_max = pt
        else:
            overall_min = XYZ(
                min(overall_min.X, pt.X),
                min(overall_min.Y, pt.Y),
                min(overall_min.Z, pt.Z)
            )
            overall_max = XYZ(
                max(overall_max.X, pt.X),
                max(overall_max.Y, pt.Y),
                max(overall_max.Z, pt.Z)
            )

if overall_min is None:
    forms.alert(
        "Could not read geometry from the selected linked elements.",
        exitscript=True
    )

# 3. Apply padding
overall_min = XYZ(overall_min.X - PADDING_FT, overall_min.Y - PADDING_FT, overall_min.Z - PADDING_FT)
overall_max = XYZ(overall_max.X + PADDING_FT, overall_max.Y + PADDING_FT, overall_max.Z + PADDING_FT)

new_box = BoundingBoxXYZ()
new_box.Min = overall_min
new_box.Max = overall_max

# 4. Find (or prepare to create) the dedicated 3D view
existing_view = None
for v in FilteredElementCollector(doc).OfClass(View3D).ToElements():
    if not v.IsTemplate and v.Name == VIEW_NAME:
        existing_view = v
        break

t = Transaction(doc, "Section Box from Linked Elements")
t.Start()
try:
    if existing_view is not None:
        target_view = existing_view
    else:
        # Find a 3D ViewFamilyType to base the new view on
        vft = None
        for candidate in FilteredElementCollector(doc).OfClass(ViewFamilyType).ToElements():
            if candidate.ViewFamily == ViewFamily.ThreeDimensional:
                vft = candidate
                break

        if vft is None:
            t.RollBack()
            forms.alert("No 3D view family type found in the project.", exitscript=True)

        target_view = View3D.CreateIsometric(doc, vft.Id)
        target_view.Name = VIEW_NAME

    target_view.IsSectionBoxActive = True
    target_view.SetSectionBox(new_box)
    t.Commit()
except Exception as ex:
    t.RollBack()
    forms.alert("Failed to set section box:\n{}".format(ex), exitscript=True)

# 5. Switch to the 3D view and zoom to fit
uidoc.ActiveView = target_view
uidoc.RefreshActiveView()
for uiview in uidoc.GetOpenUIViews():
    if uiview.ViewId == target_view.Id:
        uiview.ZoomToFit()
        break