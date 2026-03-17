import FreeCAD
import FreeCADGui
import time


def highlight_subelements(obj, subelements, doc=None):
    """
    Highlights the specified subelements (e.g., faces) of a FreeCAD object in the 3D view.
    Args:
        obj: The FreeCAD object (e.g., Part::Box).
        subelements: A list of subelement names (e.g., ["Face6", "Face2"]).
        doc: The FreeCAD document. If None, uses FreeCAD.ActiveDocument.
    """
    if doc is None:
        doc = FreeCAD.ActiveDocument
    if doc is None:
        doc = FreeCAD.newDocument("HighlightDemo")
    FreeCAD.setActiveDocument(doc.Name)
    FreeCADGui.ActiveDocument = FreeCADGui.getDocument(doc.Name)
    if obj.ViewObject is not None:
        obj.ViewObject.Visibility = True
    FreeCADGui.updateGui()
    time.sleep(0.1)
    FreeCADGui.Selection.clearSelection()
    for sub in subelements:
        FreeCADGui.Selection.addSelection(obj, sub)
