import FreeCAD
import FreeCADGui
from PySide import QtWidgets, QtCore, QtGui
import Part
import os

# Base directory of this macro (for custom icons: icons/SelectionSet.svg, icons/SelectionSetLink.svg).
_MACRO_DIR = os.path.dirname(os.path.abspath(__file__))

# Macro version: used when loading (printed in Report view) and in test runner to confirm the loaded code version.
__version__ = "1.0.0"


def _macro_icons_dir():
    """Return the absolute path to the macro's icons directory. Works when run as FCMacro or via import."""
    return os.path.abspath(os.path.join(_MACRO_DIR, "icons"))


def get_toolbar_icon(icon_name):
    """
    Return a QIcon for toolbar buttons. icon_name: 'SelectionSet' or 'SelectionSetLink'.
    Returns QIcon or None if file not found (buttons will show text only).
    """
    icons_dir = _macro_icons_dir()
    names = {
        "SelectionSet": "SelectionSet.svg",
        "SelectionSetLink": "SelectionSetLink.svg",
        "RunTest1": "RunTest1.svg",
        "RunTest2": "RunTest2.svg",
        "RunTest3": "RunTest3.svg",
        "RunTestAll": "RunTestAll.svg",
    }
    path = os.path.join(icons_dir, names.get(icon_name, icon_name))
    if os.path.isfile(path):
        return QtGui.QIcon(os.path.abspath(path))
    return None


"""
selection_set_core: SelectionSet logic and GUI (no test runner).
Create/expand/update SelectionSet, observer, ViewProvider, commands, selection logic.

FreeCAD version compatibility (1.0 and 1.1):
- This module is written to run on both FreeCAD 1.0.x and 1.1.x. Where APIs differ,
  we use hasattr() checks or try/except and keep the 1.0 code path; 1.1-specific
  behaviour is added only when needed (e.g. deprecated properties replaced in 1.1).
- Selection observer: we try both FreeCADGui.SelectionObserverPython (if present) and
  FreeCADGui.Selection.addObserver/removeObserver so that both 1.0 and 1.1 work.
"""


def _freecad_version_tuple():
    """
    Return (major, minor) from FreeCAD.Version() for version checks.
    Example: (1, 0) or (1, 1). Used to adopt 1.1 APIs while keeping 1.0 behaviour.
    """
    try:
        v = getattr(FreeCAD, "Version", None) or ()
        if isinstance(v, (list, tuple)) and len(v) >= 2:
            return (int(v[0]), int(v[1]))
        if isinstance(v, str):
            parts = v.split(".")[:2]
            return (int(parts[0]), int(parts[1])) if len(parts) >= 2 else (1, 0)
    except Exception:
        pass
    return (1, 0)


# Implementation of TODO-FC 1: GUI element for selection sets in FreeCAD

# Set to True for extra debug output in Report view (observer args, attach/detach, view updates).
# Temporarily default to True while debugging deletion behaviour; turn off for normal use later.
DEBUG_SELECTIONSET = True

# Set to False to avoid duplicate-observer TypeError (FreeCAD can keep old observers; we can't remove them).
# When False: use double-click or context menu on a SelectionSet in the tree to expand to its elements.
# When True: single-click on a SelectionSet in the tree also expands (requires restarting FreeCAD once after enabling).
USE_SELECTION_OBSERVER = False

# When False, createSelectionSetFromCurrent() does not print [DEBUG], "--- SelectionSet Button Pressed ---",
# or per-face/per-solid lines (keeps test runner output readable). Set to False during run_selectionset_tests().
SELECTIONSET_VERBOSE_CALLBACK = True

# Flag to avoid re-entrancy when we replace selection with SelectionSet elements
_selection_set_expanding = False


def _log_to_report(msg):
    """Write message to Report view if available, else to console. Avoids double output."""
    report = (
        FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")
        if FreeCADGui.getMainWindow()
        else None
    )
    if report:
        report.append(msg)
    else:
        print(msg)


def _debug_log(msg):
    """Log only when DEBUG_SELECTIONSET is True."""
    if DEBUG_SELECTIONSET:
        _log_to_report("[SelectionSet DEBUG] " + msg)


# ---------- Observer and GUI: tree click expands SelectionSet to elements ----------
class SelectionSetObserver:
    """
    Observer that reacts when a SelectionSet is selected (e.g. by click in tree view).
    Replaces the selection with the stored ElementList and logs to Report view.
    """

    def addSelection(self, *args):
        """Called when something is added to selection. FreeCAD may pass (doc, obj, sub) or (doc, obj, sub, x, y, z). Accept *args to avoid TypeError with any argument count."""
        global _selection_set_expanding
        doc_name = args[0] if len(args) > 0 else None
        obj_name = args[1] if len(args) > 1 else None
        sub_name = args[2] if len(args) > 2 else None
        _debug_log(
            "addSelection called: doc=%s obj=%s sub=%s n_args=%s"
            % (doc_name, obj_name, sub_name, len(args))
        )
        if _selection_set_expanding:
            _debug_log("  (skipped: re-entrancy guard)")
            return
        if not doc_name or not obj_name:
            return
        try:
            doc = FreeCAD.getDocument(doc_name)
            obj = doc.getObject(obj_name) if doc else None
        except Exception as e:
            _debug_log("  getDocument/getObject failed: %s" % e)
            return
        if not obj or not hasattr(obj, "ElementList"):
            return
        element_list = list(obj.ElementList) if obj.ElementList else []
        _log_to_report("\n--- [SelectionSet] Clicked '%s' in tree ---" % obj_name)
        _log_to_report(
            "  ElementList (%d items): %s" % (len(element_list), element_list)
        )
        _selection_set_expanding = True
        try:
            # Build obj -> list of subelement names (same as showSelection)
            per_obj = {}
            for entry in element_list:
                if "." in entry:
                    part_name, sub = entry.split(".", 1)
                    per_obj.setdefault(part_name, []).append(sub)
                else:
                    per_obj.setdefault(entry, [])
            # Clear selection first so we're in a known state
            FreeCADGui.Selection.clearSelection()
            # Clear showSubset on all; only touch Highlighted where supported (PartGui.ViewProviderPartExt has no Highlighted)
            for o in doc.Objects:
                try:
                    if hasattr(o, "ViewObject") and o.ViewObject:
                        vp = o.ViewObject
                        try:
                            if type(vp).__name__ != "ViewProviderPartExt" and hasattr(
                                vp, "Highlighted"
                            ):
                                vp.Highlighted = False
                        except (AttributeError, Exception):
                            pass
                        if hasattr(vp, "showSubset"):
                            vp.showSubset([])
                except Exception as ex:
                    _debug_log(
                        "  clear view for %s: %s" % (getattr(o, "Name", "?"), ex)
                    )
            # Add elements to selection first (so selection always works even if view updates fail)
            for part_name, sublist in per_obj.items():
                o = doc.getObject(part_name)
                if o:
                    for sub in sublist:
                        FreeCADGui.Selection.addSelection(o, sub)
                    if not sublist:
                        FreeCADGui.Selection.addSelection(o)
            # Set showSubset so only our subelements are highlighted (avoid whole-body highlight)
            for part_name, sublist in per_obj.items():
                o = doc.getObject(part_name)
                if o and hasattr(o, "ViewObject") and o.ViewObject and sublist:
                    try:
                        if hasattr(o.ViewObject, "showSubset"):
                            o.ViewObject.showSubset(sublist)
                            _debug_log("  showSubset(%s) on %s" % (sublist, part_name))
                        else:
                            try:
                                vp = o.ViewObject
                                if type(
                                    vp
                                ).__name__ != "ViewProviderPartExt" and hasattr(
                                    vp, "Highlighted"
                                ):
                                    vp.Highlighted = True
                            except (AttributeError, Exception):
                                pass
                    except Exception as ex:
                        _debug_log(
                            "  showSubset/Highlight for %s: %s" % (part_name, ex)
                        )
            _log_to_report(
                "  -> Selection replaced with %d elements from ElementList."
                % len(element_list)
            )
        except Exception as e:
            _log_to_report("  [SelectionSet] addSelection error: %s" % e)
            import traceback

            _log_to_report(traceback.format_exc())
        finally:
            _selection_set_expanding = False

    def removeSelection(self, *args):
        """Called when something is removed from selection. Accept *args to match any argument count."""
        pass


# Keep a reference so the observer is not garbage-collected
_selectionset_observer = None


def _detach_selectionset_observer(quiet=False):
    """Remove the SelectionSet observer. Called automatically on macro run to avoid duplicate observers. quiet=True: no report messages."""
    global _selectionset_observer
    if _selectionset_observer is None:
        if not quiet:
            _log_to_report("[SelectionSet] No observer to remove.")
        return
    try:
        # FreeCAD 1.0 vs 1.1: 1.1 may expose SelectionObserverPython; 1.0 uses Selection.removeObserver. Try both so both versions work.
        if hasattr(FreeCADGui, "SelectionObserverPython") and hasattr(
            FreeCADGui.SelectionObserverPython, "removeObserver"
        ):
            FreeCADGui.SelectionObserverPython.removeObserver(_selectionset_observer)
            _debug_log("removeObserver(SelectionObserverPython) succeeded.")
        elif hasattr(FreeCADGui.Selection, "removeObserver"):
            FreeCADGui.Selection.removeObserver(_selectionset_observer)
            _debug_log("Selection.removeObserver succeeded.")
        else:
            if not quiet:
                _log_to_report("[SelectionSet] No removeObserver API found.")
    except Exception as e:
        if not quiet:
            _log_to_report("[SelectionSet] removeObserver failed: %s" % e)
    _selectionset_observer = None
    if not quiet:
        _log_to_report(
            "[SelectionSet] Observer removed. Run the macro again to attach a fresh one."
        )


def _attach_selectionset_observer():
    """Attach the SelectionSet observer so clicking a SelectionSet in tree expands to elements."""
    global _selectionset_observer
    try:
        # Remove old observer first so re-running the macro uses the current code (avoids stale 4-arg addSelection)
        if _selectionset_observer is not None:
            _debug_log("Removing existing observer before attaching new one.")
            try:
                if hasattr(FreeCADGui, "SelectionObserverPython") and hasattr(
                    FreeCADGui.SelectionObserverPython, "removeObserver"
                ):
                    FreeCADGui.SelectionObserverPython.removeObserver(
                        _selectionset_observer
                    )
                elif hasattr(FreeCADGui.Selection, "removeObserver"):
                    FreeCADGui.Selection.removeObserver(_selectionset_observer)
            except Exception as ex:
                _debug_log("removeObserver raised: %s" % ex)
            _selectionset_observer = None
        _selectionset_observer = SelectionSetObserver()
        # FreeCAD 1.0 vs 1.1: prefer SelectionObserverPython if available (1.1), else Selection.addObserver (1.0).
        if hasattr(FreeCADGui, "SelectionObserverPython") and hasattr(
            FreeCADGui.SelectionObserverPython, "addObserver"
        ):
            FreeCADGui.SelectionObserverPython.addObserver(_selectionset_observer)
            _debug_log("SelectionObserverPython.addObserver used.")
        else:
            FreeCADGui.Selection.addObserver(_selectionset_observer)
            _debug_log("Selection.addObserver used.")
        _log_to_report(
            "[SelectionSet] Observer attached: clicking a SelectionSet in tree will expand to its elements."
        )
    except Exception as e:
        _log_to_report("[SelectionSet] Could not attach observer: %s" % e)


# ---------- ViewProvider: double-click expands selection and can transfer to FEM constraint ----------
class ViewProviderSelectionSet:
    def __init__(self, obj):
        obj.Proxy = self
        # Set immediately so tree does not show grayed (App::FeaturePython defaults to Visibility=False)
        obj.Visibility = True

    def attach(self, vobj):
        """Give the view a root node and keep Visibility=True so the object is not grayed out in the tree."""
        try:
            vobj.addProperty(
                "App::PropertyBool",
                "ShowCoordinatePoint",
                "Display",
                "Show the coordinate point in the 3D view when VolumeMode is 'coordinates'.",
            )
            vobj.ShowCoordinatePoint = True
        except Exception:
            pass
        try:
            from pivy import coin

            root = coin.SoSeparator()
            self._point_sep = (
                coin.SoSeparator()
            )  # child we update for the coordinate point marker
            root.addChild(self._point_sep)
            vobj.addDisplayMode(root, "Default")
            if hasattr(vobj, "setDisplayMode"):
                vobj.setDisplayMode("Default")
            # Set DisplayMode property so tree/view does not gray out (empty DisplayMode = no representation)
            if hasattr(vobj, "DisplayMode"):
                vobj.DisplayMode = "Default"
            if hasattr(vobj, "Object") and vobj.Object:
                self._update_coordinate_point_marker(vobj.Object)
        except Exception:
            pass
        # Set visibility last so the tree view shows the object as visible (not grayed).
        vobj.Visibility = True
        if (
            hasattr(vobj, "Object")
            and vobj.Object
            and hasattr(vobj.Object, "Visibility")
        ):
            try:
                vobj.Object.Visibility = True
            except Exception:
                pass

        # Delay setting DisplayMode so it runs after view is fully built (enum may be filled from getDisplayModes then)
        def _set_display_mode_default():
            try:
                if (
                    hasattr(vobj, "DisplayMode")
                    and getattr(vobj, "DisplayMode", "") != "Default"
                ):
                    vobj.DisplayMode = "Default"
            except Exception:
                pass

        try:
            QtCore.QTimer.singleShot(100, _set_display_mode_default)
        except Exception:
            pass

    def updateData(self, obj, prop):
        """When CoordinatePoint, VolumeMode, SelectionShape, ShowCoordinatePointMarker, or CoordinatePointSize change, update the point marker in 3D view."""
        if prop not in (
            "CoordinatePoint",
            "VolumeMode",
            "SelectionShape",
            "ShowCoordinatePointMarker",
            "CoordinatePointSize",
        ):
            return
        try:
            self._update_coordinate_point_marker(obj)
        except Exception:
            pass

    def _update_coordinate_point_marker(self, set_obj):
        """Draw or clear a small sphere at the coordinate point when VolumeMode is coordinates, show-marker is True (Data or View), and the SelectionSet is visible."""
        if not hasattr(self, "_point_sep"):
            return
        from pivy import coin

        vobj = set_obj.ViewObject if hasattr(set_obj, "ViewObject") else None
        # Data tab (1.1): ShowCoordinatePointMarker; View tab: ShowCoordinatePoint. Either False hides the sphere.
        show_data = getattr(set_obj, "ShowCoordinatePointMarker", True)
        show_view = getattr(vobj, "ShowCoordinatePoint", True) if vobj else True
        if not vobj or not show_data or not show_view:
            self._point_sep.removeAllChildren()
            return
        # Only skip drawing when explicitly hidden; default to True so point shows when visibility is toggled on
        vis = getattr(vobj, "Visibility", True)
        if hasattr(vobj, "Visibility") and vis is False:
            self._point_sep.removeAllChildren()
            return
        if getattr(set_obj, "VolumeMode", None) != "coordinates":
            self._point_sep.removeAllChildren()
            return
        point = _get_coordinate_point_from_set(set_obj)
        self._point_sep.removeAllChildren()
        if point is None:
            return
        try:
            # Marker size: use CoordinatePointSize when present; fall back to 0.5 (old default, ~1 mm diameter)
            size = getattr(set_obj, "CoordinatePointSize", 0.5)
            try:
                size = float(size)
            except Exception:
                size = 0.5
            if size <= 0:
                size = 0.5
            trans = coin.SoTranslation()
            trans.translation.setValue([point.x, point.y, point.z])
            sphere = coin.SoSphere()
            sphere.radius = size  # user‑adjustable marker size in model units
            mat = coin.SoMaterial()
            mat.diffuseColor.setValue(coin.SbColor(0.9, 0.75, 0.0))  # yellow/gold
            self._point_sep.addChild(trans)
            self._point_sep.addChild(mat)
            self._point_sep.addChild(sphere)
        except Exception:
            pass

    def onChanged(self, vobj, prop):
        """When ShowCoordinatePoint or Visibility changes, update the point marker (so un-hiding shows the point)."""
        if prop in ("ShowCoordinatePoint", "Visibility") and hasattr(vobj, "Object"):
            self._update_coordinate_point_marker(vobj.Object)

    def getDisplayModes(self, vobj):
        """Return list of display mode names so DisplayMode property is non-empty (required for show/hide and tree)."""
        return ["Default"]

    def __getstate__(self):
        """Return serializable state for save; do not store Coin (SoSeparator) objects."""
        return None

    def __setstate__(self, state):
        """Restore after load; attach() will recreate the scene graph."""
        pass

    def getIcon(self):
        """Return path to custom icon so tree view shows it when run via FCMacro or normal macro."""
        icons_dir = _macro_icons_dir()
        for name in ("SelectionSet.svg", "selection_set.svg"):
            path = os.path.join(icons_dir, name)
            if path and os.path.isfile(path):
                return os.path.abspath(path)
        return ""

    def setupContextMenu(self, obj, menu):
        """Add SelectionSet actions to the context menu (called by FreeCAD when right-clicking this object in the tree)."""
        set_obj = obj.Object if hasattr(obj, "Object") else obj
        if not set_obj or not hasattr(set_obj, "ElementList"):
            return
        menu.addSeparator()
        # Submenu for all SelectionSet-related actions; standard submenu so name and items always work
        submenu = QtWidgets.QMenu("SelectionSet", menu)
        submenu.addAction(
            "Create SelectionSetLink",
            lambda s=set_obj: _create_link_from_selectionset_menu(s),
        )
        submenu.addSeparator()
        submenu.addAction(
            "Update SelectionSet", lambda: _update_selectionset_from_menu(set_obj)
        )
        submenu.addAction(
            "Delete SelectionSet only (keep shapes)",
            lambda: _delete_selectionset_only(set_obj),
        )
        submenu.addSeparator()
        submenu.addAction(
            "Expand to elements (replace selection)",
            lambda: selectElementsFromSet(set_obj),
        )
        submenu.addAction(
            "Add elements from list to selection",
            lambda: _add_elements_from_set_to_selection(set_obj, clear_first=False),
        )
        submenu.addSeparator()
        submenu.addAction(
            "Add main shape to selection",
            lambda: _add_main_shape_to_selection(set_obj, clear_first=False),
        )
        submenu.addAction(
            "Add selection shape to selection",
            lambda: _add_selection_shape_to_selection(set_obj, clear_first=False),
        )
        submenu.addAction(
            "Add both shapes to selection",
            lambda: _add_both_shapes_to_selection(set_obj, clear_first=False),
        )
        submenu.addSeparator()
        submenu.addAction(
            "Select only main shape",
            lambda: _add_main_shape_to_selection(set_obj, clear_first=True),
        )
        submenu.addAction(
            "Select only selection shape",
            lambda: _add_selection_shape_to_selection(set_obj, clear_first=True),
        )
        submenu.addAction(
            "Select only elements",
            lambda: _add_elements_from_set_to_selection(set_obj, clear_first=True),
        )
        # Make the "SelectionSet" entry in the parent menu bold and optionally tinted
        action = submenu.menuAction()
        font = action.font()
        font.setBold(True)
        action.setFont(font)
        action.setObjectName("SelectionSetSubmenuTrigger")
        try:
            menu.setStyleSheet(
                (menu.styleSheet() or "")
                + " QMenu::item#SelectionSetSubmenuTrigger { background-color: #e0e8f0; } "
            )
        except Exception:
            pass
        menu.addMenu(submenu)

    def onSelection(self, vp):
        # Debug: log immediately so we see something when SelectionSet is selected
        _log_to_report("\n--- [SelectionSet] ViewProvider.onSelection called ---")
        selection_set = vp.Object
        element_list = (
            list(selection_set.ElementList) if selection_set.ElementList else []
        )
        _log_to_report(
            "  SelectionSet '%s' selected. ElementList (%d items): %s"
            % (selection_set.Name, len(element_list), element_list)
        )

        import FreeCAD
        from highlight_subelements import highlight_subelements

        # Remove the SelectionSet object from the selection (so we don't keep it selected)
        FreeCADGui.Selection.removeSelection(selection_set)

        # Parse element_list to build a mapping of objects and their subelements
        obj_map = {}
        for elem in element_list:
            if "." in elem:
                part_name, sub_name = elem.split(".", 1)
                _log_to_report(
                    "  Processing: %s -> part: %s, subelement: %s"
                    % (elem, part_name, sub_name)
                )
                obj = FreeCAD.ActiveDocument.getObject(part_name)
                if obj:
                    obj_map.setdefault(obj, []).append(sub_name)
                else:
                    FreeCAD.Console.PrintError(
                        "Object %s not found in document!\n" % part_name
                    )

        # Add all elements to the selection (so they appear in the selection and property panel)
        for obj, subelements in obj_map.items():
            for sub in subelements:
                FreeCADGui.Selection.addSelection(obj, sub)
            if not subelements:
                FreeCADGui.Selection.addSelection(obj)
        _log_to_report("  -> Added %d elements to selection." % len(element_list))

        # Highlight all subelements in the 3D view
        for obj, subelements in obj_map.items():
            if subelements:
                highlight_subelements(obj, subelements, FreeCAD.ActiveDocument)

        # If a FEM constraint is selected, set its References property.
        # FEM expects a flat list of (object, subelement_name) per reference, e.g. [(obj, "Face6"), (obj, "Solid1")].
        sel = FreeCADGui.Selection.getSelection()
        refs_flat = [(obj, sub) for obj, sublist in obj_map.items() for sub in sublist]
        if sel and hasattr(sel[0], "References"):
            constraint = sel[0]
            if refs_flat:
                constraint.References = refs_flat
                constraint.touch()  # ensure constraint and task panel refresh
                if hasattr(constraint, "Document") and constraint.Document:
                    constraint.Document.recompute()
            _log_to_report(
                "  -> Transferred %d refs to FEM constraint References (flat format for Face/Solid)."
                % len(refs_flat)
            )
        return True


# ---------- SelectionSet FeaturePython object (properties, showSelection) ----------


def _update_selection_property_editor_modes(obj):
    """
    Hide or show selection-related properties in the Data tab depending on Mode/VolumeMode.

    - CoordinatePoint, ShowCoordinatePointMarker, CoordinatePointSize: only shown when VolumeMode == 'coordinates'.
    - SolidFilterRefs: only shown when Mode == 'faces' (it filters faces by solid).
    """
    try:
        volume_mode = getattr(obj, "VolumeMode", None) or "fully_inside"
        mode = getattr(obj, "Mode", None) or "faces"
        # CoordinatePoint, ShowCoordinatePointMarker, CoordinatePointSize visible only for coordinates mode
        for prop in ("CoordinatePoint", "ShowCoordinatePointMarker", "CoordinatePointSize"):
            if hasattr(obj, prop) and hasattr(obj, "setEditorMode"):
                obj.setEditorMode(
                    prop, 0 if volume_mode == "coordinates" else 2
                )  # 0 = normal, 2 = hidden
        # SolidFilterRefs only relevant for faces-based selections
        if hasattr(obj, "SolidFilterRefs") and hasattr(obj, "setEditorMode"):
            obj.setEditorMode(
                "SolidFilterRefs",
                0 if mode == "faces" else 2,
            )
    except Exception:
        # Editor-mode customization is best-effort only; do not disturb core logic.
        pass


class SelectionSet:
    def __init__(self, obj):
        obj.addProperty(
            "App::PropertyStringList", "ElementList", "Selection", "List of subelements"
        )
        obj.addProperty(
            "App::PropertyLink",
            "MainObject",
            "Shapes",
            "Main object (whose faces/solids are filtered by the selection shape)",
        )
        obj.addProperty(
            "App::PropertyLink",
            "SelectionShape",
            "Shapes",
            "Selection-defining shape (volume filter)",
        )
        obj.addProperty(
            "App::PropertyEnumeration",
            "Mode",
            "Selection",
            "Selection mode: faces or solids",
        )
        obj.Mode = ["faces", "solids"]
        obj.Mode = "faces"
        obj.addProperty(
            "App::PropertyEnumeration",
            "VolumeMode",
            "Selection",
            "Volume/coordinates filter: fully_inside, intersection, or coordinates (point-in-face/solid)",
        )
        obj.VolumeMode = ["fully_inside", "intersection", "coordinates"]
        obj.VolumeMode = "fully_inside"
        obj.addProperty(
            "App::PropertyVector",
            "CoordinatePoint",
            "Selection",
            "Point (x,y,z) for coordinates mode. If (0,0,0) and SelectionShape is set, shape center is used.",
        )
        try:
            obj.addProperty(
                "App::PropertyBool",
                "ShowCoordinatePointMarker",
                "Display",
                "Show the coordinate point (sphere) in the 3D view when VolumeMode is 'coordinates'. Uncheck to hide.",
            )
            obj.ShowCoordinatePointMarker = True
        except Exception:
            pass
        try:
            obj.addProperty(
                "App::PropertyFloat",
                "CoordinatePointSize",
                "Display",
                "Radius of the coordinate point marker (sphere) in model units. Similar to FEM fixed/force symbol size.",
            )
            obj.CoordinatePointSize = 0.5  # keep existing visual default (~1 mm diameter on most models)
        except Exception:
            pass
        try:
            obj.addProperty(
                "App::PropertyLinkSubList",
                "SolidFilterRefs",
                "Selection",
                "Solids to filter by: click ... to select solids (like Main Object / Selection Shape). Only faces from these solids are included.",
            )
        except Exception:
            pass
        # Meta: macro version when this object was created (read-only)
        try:
            obj.addProperty(
                "App::PropertyString",
                "MacroVersion",
                "Meta",
                "SelectionSet macro version that created this object.",
            )
            obj.MacroVersion = __version__
            obj.setEditorMode("MacroVersion", 1)  # ReadOnly
        except Exception:
            pass
        # Apply initial editor-mode rules so only relevant fields show up
        _update_selection_property_editor_modes(obj)
        obj.Proxy = self
        obj.ViewObject.Proxy = ViewProviderSelectionSet(obj.ViewObject)

    def onDocumentRestored(self, obj):
        # Backward compatibility: add new properties if missing (old documents)
        if not hasattr(obj, "MainObject"):
            obj.addProperty(
                "App::PropertyLink",
                "MainObject",
                "Shapes",
                "Main object (whose faces/solids are filtered by the selection shape)",
            )
        if not hasattr(obj, "SelectionShape"):
            obj.addProperty(
                "App::PropertyLink",
                "SelectionShape",
                "Shapes",
                "Selection-defining shape (volume filter)",
            )
        if not hasattr(obj, "Mode"):
            obj.addProperty(
                "App::PropertyEnumeration",
                "Mode",
                "Selection",
                "Selection mode: faces or solids",
            )
            obj.Mode = ["faces", "solids"]
            obj.Mode = "faces"
        if not hasattr(obj, "VolumeMode"):
            obj.addProperty(
                "App::PropertyEnumeration",
                "VolumeMode",
                "Selection",
                "Volume/coordinates filter: fully_inside, intersection, or coordinates (point-in-face/solid)",
            )
            obj.VolumeMode = ["fully_inside", "intersection", "coordinates"]
            obj.VolumeMode = "fully_inside"
        if not hasattr(obj, "CoordinatePoint"):
            obj.addProperty(
                "App::PropertyVector",
                "CoordinatePoint",
                "Selection",
                "Point (x,y,z) for coordinates mode. If (0,0,0) and SelectionShape is set, shape center is used.",
            )
        if not hasattr(obj, "ShowCoordinatePointMarker"):
            try:
                obj.addProperty(
                    "App::PropertyBool",
                    "ShowCoordinatePointMarker",
                    "Display",
                    "Show the coordinate point (sphere) in the 3D view when VolumeMode is 'coordinates'. Uncheck to hide.",
                )
                obj.ShowCoordinatePointMarker = True
            except Exception:
                pass
        if not hasattr(obj, "SolidFilterRefs"):
            try:
                obj.addProperty(
                    "App::PropertyLinkSubList",
                    "SolidFilterRefs",
                    "Selection",
                    "Solids to filter by: click ... to select solids (like Main Object / Selection Shape). Only faces from these solids are included.",
                )
            except Exception:
                pass
        if not hasattr(obj, "MacroVersion"):
            try:
                obj.addProperty(
                    "App::PropertyString",
                    "MacroVersion",
                    "Meta",
                    "SelectionSet macro version that created this object.",
                )
                obj.MacroVersion = __version__
                obj.setEditorMode("MacroVersion", 1)
            except Exception:
                pass
        # Migrate old SolidFilterList (string list) to SolidFilterRefs if present and refs empty
        if getattr(obj, "SolidFilterRefs", None) is not None:
            existing = obj.SolidFilterRefs or []
            if (not existing or len(existing) == 0) and getattr(
                obj, "SolidFilterList", None
            ):
                mig = []
                doc = getattr(obj, "Document", None) or FreeCAD.ActiveDocument
                for s in obj.SolidFilterList or []:
                    s = (s or "").strip()
                    if ".Solid" in s and "." in s:
                        name, sub = s.split(".", 1)
                        o = doc.getObject(name) if doc else None
                        if o and sub:
                            mig.append((o, sub))
                if mig:
                    obj.SolidFilterRefs = mig
        # Update editor modes after restoring/migrating properties
        _update_selection_property_editor_modes(obj)
        obj.Proxy = self
        obj.ViewObject.Proxy = ViewProviderSelectionSet(obj.ViewObject)

    def onChanged(self, obj, prop):
        """
        React to Mode / VolumeMode changes to keep the Selection group clean:
        only show CoordinatePoint for coordinates mode, and SolidFilterRefs for faces mode.
        """
        if prop in ("Mode", "VolumeMode"):
            _update_selection_property_editor_modes(obj)

    def onDelete(self, obj, subelements):
        """
        When a SelectionSet is deleted, remove it from any SelectionSetLink AddSelectionSets/SubtractSelectionSets
        so that links are effectively 'unlinked' instead of holding dangling references. Also clear MainObject
        and SelectionShape links so that deleting the SelectionSet does not cause its main/selection shapes to
        be treated as dependants to delete. Links and shapes are not deleted; their inputs are just cleaned up.
        """
        try:
            FreeCAD.Console.PrintMessage(
                "[SelectionSet] onDelete for '%s' (MainObject=%s, SelectionShape=%s)\n"
                % (
                    obj.Name,
                    getattr(obj, "MainObject", None),
                    getattr(obj, "SelectionShape", None),
                )
            )
        except Exception:
            pass
        doc = getattr(obj, "Document", None) or FreeCAD.ActiveDocument
        if not doc:
            return True
        # First, break direct links to shapes so they are not considered as dependants
        try:
            if hasattr(obj, "MainObject"):
                FreeCAD.Console.PrintMessage(
                    "[SelectionSet] onDelete: clearing MainObject of '%s' (was %s)\n"
                    % (obj.Name, getattr(obj, "MainObject", None))
                )
                obj.MainObject = None
            if hasattr(obj, "SelectionShape"):
                FreeCAD.Console.PrintMessage(
                    "[SelectionSet] onDelete: clearing SelectionShape of '%s' (was %s)\n"
                    % (obj.Name, getattr(obj, "SelectionShape", None))
                )
                obj.SelectionShape = None
        except Exception:
            pass
        try:
            for o in doc.Objects:
                proxy = getattr(o, "Proxy", None)
                if isinstance(proxy, SelectionSetLink):
                    try:
                        add_sets = list(getattr(o, "AddSelectionSets", []) or [])
                        sub_sets = list(getattr(o, "SubtractSelectionSets", []) or [])
                        if obj in add_sets:
                            FreeCAD.Console.PrintMessage(
                                "[SelectionSet] onDelete: removing '%s' from AddSelectionSets of link '%s'\n"
                                % (obj.Name, o.Name)
                            )
                            add_sets = [s for s in add_sets if s is not obj]
                            o.AddSelectionSets = add_sets
                        if obj in sub_sets:
                            FreeCAD.Console.PrintMessage(
                                "[SelectionSet] onDelete: removing '%s' from SubtractSelectionSets of link '%s'\n"
                                % (obj.Name, o.Name)
                            )
                            sub_sets = [s for s in sub_sets if s is not obj]
                            o.SubtractSelectionSets = sub_sets
                    except Exception:
                        pass
        except Exception:
            # Do not block deletion if cleanup fails
            return True
        return True

    def showSelection(self, obj):
        FreeCADGui.Selection.clearSelection()
        for o in FreeCAD.ActiveDocument.Objects:
            if hasattr(o, "ViewObject") and o.ViewObject:
                vp = o.ViewObject
                try:
                    if type(vp).__name__ != "ViewProviderPartExt" and hasattr(
                        vp, "Highlighted"
                    ):
                        vp.Highlighted = False
                except (AttributeError, Exception):
                    pass
                if hasattr(vp, "showSubset"):
                    vp.showSubset([])
        per_obj = {}
        for entry in obj.ElementList:
            if "." in entry:
                obj_name, sub = entry.split(".", 1)
                per_obj.setdefault(obj_name, []).append(sub)
            else:
                per_obj.setdefault(entry, [])
        for obj_name, sublist in per_obj.items():
            fc_obj = FreeCAD.ActiveDocument.getObject(obj_name)
            if fc_obj:
                if hasattr(fc_obj.ViewObject, "showSubset") and sublist:
                    fc_obj.ViewObject.showSubset(sublist)
                else:
                    try:
                        vp = fc_obj.ViewObject
                        if type(vp).__name__ != "ViewProviderPartExt" and hasattr(
                            vp, "Highlighted"
                        ):
                            vp.Highlighted = True
                    except (AttributeError, Exception):
                        pass
                for sub in sublist:
                    FreeCADGui.Selection.addSelection(fc_obj, sub)
                if not sublist:
                    FreeCADGui.Selection.addSelection(fc_obj)


# ---------- Commands: HighlightSelectionSet, UpdateSelectionSet ----------
class SelectionSetHighlightCommand:
    def GetResources(self):
        return {
            "MenuText": "Highlight SelectionSet",
            "ToolTip": "Restore and highlight elements stored in SelectionSet.",
        }

    def Activated(self):
        sel = FreeCADGui.Selection.getSelection()
        for obj in sel:
            if hasattr(obj, "ElementList"):
                selectElementsFromSet(obj)

    def IsActive(self):
        return True


# Register the command
FreeCADGui.addCommand("HighlightSelectionSet", SelectionSetHighlightCommand())


class UpdateSelectionSetCommand:
    """Recompute ElementList from MainObject and SelectionShape for selected SelectionSet(s)."""

    def GetResources(self):
        return {
            "MenuText": "Update SelectionSet",
            "ToolTip": "Recompute element list from the stored Main Object and Selection Shape. Select one or more SelectionSets in the tree first.",
        }

    def Activated(self):
        sel = FreeCADGui.Selection.getSelection()
        updated = 0
        for obj in sel:
            if hasattr(obj, "ElementList") and hasattr(obj, "MainObject"):
                if recomputeSelectionSetFromShapes(obj):
                    updated += 1
        if updated:
            print("Updated %d SelectionSet(s) from their stored shapes." % updated)

    def IsActive(self):
        return True


FreeCADGui.addCommand("UpdateSelectionSet", UpdateSelectionSetCommand())


# ---------- Tree: double-click to expand SelectionSet; Shift/Alt/Ctrl+Click for main/shape/elements ----------
_trees_selectionset_connected = set()
# Modifiers from last left-click with Shift/Alt/Ctrl (used by timer to select main/shape/elements)
_pending_selectionset_click_modifiers = None
_selectionset_tree_event_filter = (
    None  # keep reference so filter is not garbage-collected
)


def _on_selectionset_modifier_click_timeout():
    """Run after a click with Shift/Alt/Ctrl on tree: if selection is a single SelectionSet, select main/shape/elements."""
    global _pending_selectionset_click_modifiers
    mods = _pending_selectionset_click_modifiers
    _pending_selectionset_click_modifiers = None
    if mods is None:
        return
    sel = FreeCADGui.Selection.getSelection()
    if len(sel) != 1:
        return
    obj = sel[0]
    qt = QtCore.Qt
    # SelectionSet: Shift=main, Alt=shape, Ctrl=elements
    if hasattr(obj, "ElementList"):
        if mods & qt.ControlModifier:
            selectElementsFromSet(obj)
            return
        if mods & qt.AltModifier:
            _add_selection_shape_to_selection(obj, clear_first=True)
            return
        if mods & qt.ShiftModifier:
            _add_main_shape_to_selection(obj, clear_first=True)
            return
    # SelectionSetLink: Ctrl+click expands to elements
    if (
        mods & qt.ControlModifier
        and hasattr(obj, "Proxy")
        and hasattr(getattr(obj, "Proxy", None), "get_combined_refs")
    ):
        _link_expand_to_selection(obj, clear_first=True)


class _SelectionSetTreeEventFilter(QtCore.QObject):
    """Event filter to detect Shift/Alt/Ctrl+Click on tree; schedules selection of main/shape/elements."""

    def eventFilter(self, watched, event):
        global _pending_selectionset_click_modifiers
        try:
            if (
                event.type() == QtCore.QEvent.MouseButtonPress
                and event.button() == QtCore.Qt.LeftButton
            ):
                mods = event.modifiers()
                qt = QtCore.Qt
                if mods & (qt.ShiftModifier | qt.AltModifier | qt.ControlModifier):
                    _pending_selectionset_click_modifiers = mods
                    QtCore.QTimer.singleShot(
                        50, _on_selectionset_modifier_click_timeout
                    )
        except Exception:
            pass
        return False  # do not consume the event


def add_selectionset_context_menu():
    """Connect double-click so expanding a SelectionSet works without the selection observer. Does NOT set CustomContextMenu so the default FreeCAD tree context menu keeps working for all objects."""
    mw = FreeCADGui.getMainWindow()
    if not mw:
        return False

    def run_expand_for_selection():
        """If current selection contains a SelectionSet (has ElementList), expand it to stored elements."""
        sel = FreeCADGui.Selection.getSelection()
        for obj in sel:
            if hasattr(obj, "ElementList"):
                selectElementsFromSet(obj)
                return

    def double_click_event(index):
        QtCore.QTimer.singleShot(10, run_expand_for_selection)

    global _selectionset_tree_event_filter
    connected = False
    for tree in mw.findChildren(QtWidgets.QTreeView):
        if id(tree) in _trees_selectionset_connected:
            continue
        try:
            tree.doubleClicked.connect(double_click_event)
            if _selectionset_tree_event_filter is None:
                _selectionset_tree_event_filter = _SelectionSetTreeEventFilter()
            viewport = tree.viewport()
            if viewport:
                viewport.installEventFilter(_selectionset_tree_event_filter)
            _trees_selectionset_connected.add(id(tree))
            connected = True
        except Exception:
            pass
    return connected


def _connect_selectionset_tree_handlers():
    """Connect double-click/context menu for SelectionSet expand; retry once if tree not ready at load."""
    if add_selectionset_context_menu():
        return
    QtCore.QTimer.singleShot(500, add_selectionset_context_menu)


_connect_selectionset_tree_handlers()


# ---------- Selection logic: point/solid inside shape, face filtering ----------
def _point_inside_shape(point, shape, tol):
    """
    Return True if point is inside shape. Whenever the shape has Solids (single or compound,
    e.g. BooleanFragments or PartDesign Body), check inside *each* solid and return True if
    inside any, since shape.isInside() on a compound can be unreliable.
    """
    try:
        solids = getattr(shape, "Solids", None)
        if solids and len(solids) >= 1:
            return any(s.isInside(point, tol, True) for s in solids)
        return shape.isInside(point, tol, True)
    except Exception:
        return False


def _solid_fully_inside_shape(solid, shape, tol_volume=1e-9):
    """
    Return True if the solid is fully inside the shape (volume containment).

    Uses OCCT/FreeCAD boolean: solid is inside shape iff Common(solid, shape) has the same
    volume as solid (every point of solid belongs to shape). Falls back to vertex + center-of-mass
    check if the boolean operation fails or is invalid.
    tol_volume: absolute volume tolerance for equality (mm³); also scaled by max(solid.Volume, 1.)
    """
    try:
        common = solid.common(shape)
        if common.isNull():
            raise ValueError("Common is null")
        vol_solid = solid.Volume
        vol_common = common.Volume
        ref = max(vol_solid, 1.0)
        if abs(vol_common - vol_solid) <= tol_volume * ref:
            return True
        return False
    except Exception:
        # Fallback: all vertices and center of mass inside (point-in-solid test)
        all_inside = all(
            _point_inside_shape(v.Point, shape, 1e-7) for v in solid.Vertexes
        )
        center_inside = _point_inside_shape(solid.CenterOfMass, shape, 1e-7)
        return all_inside and center_inside


def _solid_intersects_shape(solid, shape, tol_volume=1e-9):
    """
    Return True if the solid has any volume intersection with the shape.

    Uses OCCT/FreeCAD boolean: Common(solid, shape).Volume > tol_volume.
    Falls back to center-of-mass or vertex inside shape if the boolean fails.
    """
    try:
        common = solid.common(shape)
        if common.isNull():
            return False
        if common.Volume > tol_volume:
            return True
        return False
    except Exception:
        # Fallback: center of mass or any vertex inside shape
        if _point_inside_shape(solid.CenterOfMass, shape, 1e-7):
            return True
        return any(_point_inside_shape(v.Point, shape, 1e-7) for v in solid.Vertexes)


def _get_coordinate_point_from_set(set_obj, tol_zero=1e-10):
    """
    Get the (x,y,z) point for coordinates mode from a SelectionSet.
    Uses CoordinatePoint if it is not (0,0,0); otherwise uses the center of SelectionShape
    if set. Returns None if no point can be determined.
    """
    try:
        v = getattr(set_obj, "CoordinatePoint", None)
        if v is not None:
            x, y, z = v.x, v.y, v.z
            if abs(x) > tol_zero or abs(y) > tol_zero or abs(z) > tol_zero:
                return FreeCAD.Vector(x, y, z)
        shape_obj = getattr(set_obj, "SelectionShape", None)
        if shape_obj and getattr(shape_obj, "Shape", None):
            s = shape_obj.Shape
            if hasattr(s, "CenterOfMass"):
                return s.CenterOfMass
            if hasattr(s, "BoundBox") and s.BoundBox:
                bb = s.BoundBox
                return FreeCAD.Vector(
                    0.5 * (bb.XMin + bb.XMax),
                    0.5 * (bb.YMin + bb.YMax),
                    0.5 * (bb.ZMin + bb.ZMax),
                )
    except Exception:
        pass
    return None


def _point_on_face(face, point, tol=1e-6):
    """Return True if point is on the face (within tolerance). Uses distToShape."""
    try:
        vertex = Part.Vertex(point)
        d = face.distToShape(vertex)
        if d and len(d) >= 1:
            val = d[0]
            dist = (
                val.Length
                if hasattr(val, "Length")
                else (float(val) if isinstance(val, (int, float)) else 1e99)
            )
            return dist <= tol
        return False
    except Exception:
        return False


def get_faces_by_point(obj, point, tol=1e-6, solid_filter_list=None):
    """
    Return list of face element names (e.g. Obj.Face3) of obj that contain the given point.
    point: FreeCAD.Vector. Uses distToShape(face, vertex) <= tol to decide.
    solid_filter_list: optional; if set, only include faces that belong to one of these solids.
    """
    point = FreeCAD.Vector(point.x, point.y, point.z)
    faces = []
    shape = getattr(obj, "Shape", None)
    if not shape or not getattr(shape, "Faces", None):
        return faces
    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")
    for i, face in enumerate(shape.Faces):
        if _point_on_face(face, point, tol):
            faces.append(f"{obj.Name}.Face{i + 1}")
            if SELECTIONSET_VERBOSE_CALLBACK and report:
                report.append("Coordinates mode: point on Face %d" % (i + 1))
    if solid_filter_list:
        solid_set = set(
            s.strip() for s in solid_filter_list if s and ".Solid" in str(s)
        )
        if solid_set:
            filtered = []
            for fe in faces:
                solid_name = _get_solid_for_face(obj, fe)
                if solid_name and solid_name in solid_set:
                    filtered.append(fe)
            faces = filtered
    return faces


def get_solids_by_point(obj, point, tol=1e-6):
    """
    Return list of solid element names (e.g. Obj.Solid2) of obj that contain the given point.
    """
    point = FreeCAD.Vector(point.x, point.y, point.z)
    solids = []
    shape = getattr(obj, "Shape", None)
    if not shape or not getattr(shape, "Solids", None):
        return solids
    for i, solid in enumerate(shape.Solids):
        try:
            if solid.isInside(point, tol, True):
                solids.append(f"{obj.Name}.Solid{i + 1}")
        except Exception:
            pass
    return solids


def _normalize_link_sub(item):
    """Get (object, subname_string) from a PropertyLinkSubList item. SubName may be tuple e.g. ('Solid3',)."""
    if (
        isinstance(item, (list, tuple))
        and len(item) >= 2
        and item[0] is not None
        and item[1] is not None
    ):
        sub = item[1]
        if isinstance(sub, (list, tuple)):
            sub = sub[0] if sub else ""
        return (item[0], str(sub).strip())
    if hasattr(item, "Object") and getattr(item, "Object", None) is not None:
        sub = getattr(item, "SubName", None) or ""
        if isinstance(sub, (list, tuple)):
            sub = sub[0] if sub else ""
        return (item.Object, str(sub).strip())
    return (None, "")


def _get_solid_filter_list_from_set(set_obj):
    """
    Return solid filter as list of strings (e.g. ['Obj.Solid1', 'Obj.Solid3']) from SolidFilterRefs.
    FreeCAD may store SubName as tuple ('Solid3',); normalize so we get 'Obj.Solid3' not 'Obj.('Solid3',)'.
    """
    refs = getattr(set_obj, "SolidFilterRefs", None) or []
    if not refs:
        return []
    result = []
    for item in refs:
        obj, sub = _normalize_link_sub(item)
        if obj and sub:
            result.append("%s.%s" % (obj.Name, sub))
    return result


def _face_to_solid_ranges(obj):
    """
    Build a mapping from 1-based face index to solid name (e.g. "Obj.Solid3").
    In Part, shape.Faces for a compound is often the concatenation of each solid's faces in order.
    Only use this if total face count matches sum of solid face counts; otherwise return empty dict.
    """
    out = {}
    shape = obj.Shape
    if not hasattr(shape, "Faces") or not hasattr(shape, "Solids") or not shape.Solids:
        return out
    total = len(shape.Faces)
    idx = 1
    for i, solid in enumerate(shape.Solids):
        n = len(solid.Faces) if hasattr(solid, "Faces") else 0
        solid_name = f"{obj.Name}.Solid{i + 1}"
        for _ in range(n):
            out[idx] = solid_name
            idx += 1
    if idx - 1 != total:
        return {}  # compound face order doesn't match solid order, use point test instead
    return out


def _get_solid_for_face(obj, face_element_name):
    """
    Return the solid element name (e.g. "Obj.Solid3") that contains the given face element.
    face_element_name must be like "ObjName.Face4". Uses point-in-solid with a majority vote
    over face vertices and center so boundary points and BooleanFragments ordering are handled.
    """
    if not face_element_name or "." not in face_element_name:
        return None
    obj_name, sub = face_element_name.split(".", 1)
    if not sub.startswith("Face"):
        return None
    try:
        face_index = int(sub[4:])  # "Face4" -> 4
    except ValueError:
        return None
    if obj_name != obj.Name:
        return None
    shape = obj.Shape
    if not hasattr(shape, "Faces") or not hasattr(shape, "Solids") or not shape.Solids:
        return None
    if face_index < 1 or face_index > len(shape.Faces):
        return None
    face = shape.Faces[face_index - 1]
    try:
        points_to_try = []
        if hasattr(face, "Vertexes") and face.Vertexes:
            points_to_try.extend([v.Point for v in face.Vertexes])
        if hasattr(face, "CenterOfMass"):
            points_to_try.append(face.CenterOfMass)
        if not points_to_try:
            return None
        # Try increasing tolerances; BooleanFragments/cut faces often need looser tol for isInside
        for tol in (1e-3, 1e-2, 1e-1):
            # Majority vote: count how many points fall inside each solid; return solid with max count.
            # When a point is inside multiple solids (boundary), assign to highest-indexed solid
            # so we prefer e.g. Solid3 over Solid2 (compound order often puts "fully inside" half last).
            counts = [0] * len(shape.Solids)
            for pt in points_to_try:
                for i in range(len(shape.Solids) - 1, -1, -1):
                    if _point_inside_shape(pt, shape.Solids[i], tol):
                        counts[i] += 1
                        break
            best = max(range(len(counts)), key=lambda i: counts[i])
            if counts[best] > 0:
                return f"{obj.Name}.Solid{best + 1}"
    except Exception:
        pass
    return None


def get_faces_by_selection_shape(
    obj, shape, tol=1e-7, volume_mode="fully_inside", solid_filter_list=None
):
    """
    Returns faces of obj filtered by the selection shape.

    volume_mode:
      - "fully_inside": include only faces that are fully inside: all vertices AND the face center
        of mass must be inside the selection shape (excludes e.g. a dome face whose vertices are
        inside but center bulges out, like the outer spherical face of the second half-sphere).
      - "intersection": include faces that have non-empty geometric intersection with the shape.
    solid_filter_list: optional list of solid element names (e.g. ["Obj.Solid3"]). If non-empty,
      only faces that belong to one of these solids are kept. Use to restrict to one half-sphere etc.
    Adds debug output for each face.
    """
    if volume_mode not in ("fully_inside", "intersection"):
        raise ValueError("volume_mode must be 'fully_inside' or 'intersection'")
    # Use tolerance so points on/near planar boundaries of selection body count as inside
    tol_inside = max(tol, 1e-4)
    faces = []
    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")
    for i, face in enumerate(obj.Shape.Faces):
        c = face.CenterOfMass
        all_inside = all(
            _point_inside_shape(v.Point, shape, tol_inside) for v in face.Vertexes
        )
        inside = _point_inside_shape(c, shape, tol_inside)
        intersected = False
        intersection_has_area = False
        try:
            common = face.common(shape)
            if not common.isNull():
                # Only count as intersection if the common has positive area (excludes point/edge contact)
                area = getattr(common, "Area", 0) or 0
                intersection_has_area = area > 1e-10
                intersected = intersection_has_area
        except Exception:
            pass
        msg = f"Face {i + 1}: center={c}, isInside(center)={inside}, all_vertices_inside={all_inside}, intersects={intersected}"
        if SELECTIONSET_VERBOSE_CALLBACK:
            if report:
                report.append(msg)
            else:
                print(msg)
        if volume_mode == "fully_inside":
            include = all_inside and inside
        else:
            include = intersected
        if include:
            faces.append(f"{obj.Name}.Face{i + 1}")
    # Filter by solid if requested (e.g. only faces of one half-sphere)
    if solid_filter_list:
        solid_set = set(
            s.strip() for s in solid_filter_list if s and ".Solid" in str(s)
        )
        if solid_set:
            filtered = []
            for fe in faces:
                solid_name = _get_solid_for_face(obj, fe)
                if solid_name and solid_name in solid_set:
                    filtered.append(fe)
            faces = filtered
    return faces


# ---------- Apply view defaults (DisplayMode, Visibility) so tree can show/hide ----------
def apply_selectionset_view_defaults(obj):
    """
    Set ViewObject DisplayMode and Visibility so the object is not grayed and can be shown/hidden.
    Call after creating a SelectionSet/SelectionSetLink (e.g. when creating from script/tests).
    """
    if obj is None:
        return
    vo = getattr(obj, "ViewObject", None)
    if vo is None:
        return
    try:
        if hasattr(vo, "DisplayMode"):
            vo.DisplayMode = "Default"
    except Exception:
        pass
    try:
        if hasattr(vo, "Visibility"):
            vo.Visibility = True
    except Exception:
        pass
    try:
        if hasattr(obj, "Visibility"):
            obj.Visibility = True
    except Exception:
        pass


def _is_selectionset(obj):
    """True if obj is a SelectionSet (our FeaturePython with ElementList, MainObject, Mode)."""
    if obj is None:
        return False
    try:
        return (
            hasattr(obj, "ElementList")
            and hasattr(obj, "MainObject")
            and hasattr(obj, "SelectionShape")
            and getattr(obj, "Mode", None) in ("faces", "solids")
        )
    except Exception:
        return False


def _is_solid_filter_selectionset(obj):
    """True if obj is a SelectionSet whose ElementList is only solids (used as solid filter source, not to update)."""
    if not _is_selectionset(obj):
        return False
    try:
        el = getattr(obj, "ElementList", None) or []
        return len(el) > 0 and all(".Solid" in str(e) for e in el)
    except Exception:
        return False


def create_or_update_selectionset_from_current(
    selection_set_name=None, force_mode=None, volume_mode=None
):
    """
    Single entry for the Create/Update SelectionSet button.
    If the current selection contains only SelectionSet objects (one or more), recompute those from
    their stored MainObject/SelectionShape. Otherwise create or update from selection (createSelectionSetFromCurrent).
    """
    sel = FreeCADGui.Selection.getSelection()
    selection_sets = [o for o in sel if _is_selectionset(o)]
    if len(selection_sets) > 0 and len(selection_sets) == len(sel):
        for set_obj in selection_sets:
            recomputeSelectionSetFromShapes(set_obj)
            apply_selectionset_view_defaults(set_obj)
        print(
            "Updated %d SelectionSet(s) from their stored shapes." % len(selection_sets)
        )
        return
    createSelectionSetFromCurrent(
        selection_set_name=selection_set_name,
        force_mode=force_mode,
        volume_mode=volume_mode,
    )


# ---------- Create and update SelectionSet from current selection ----------
def createSelectionSetFromCurrent(
    selection_set_name=None, force_mode=None, volume_mode=None
):
    """
    Create or update SelectionSet(s) from current GUI selection.
    If one or more SelectionSet objects are in the selection, those are updated (MainObject, SelectionShape,
    Mode, VolumeMode, SolidFilterRefs from the other selected objects) and recomputed; no new set is created.
    Otherwise creates a new SelectionSet.
    volume_mode: for faces with 2 objects, 'fully_inside' or 'intersection'. Default when None: 'fully_inside'.
    """
    # if selection_set_name is provided, use it; otherwise generate a default name (Name; Label set to "SelectionSet ..." below)
    if selection_set_name is None:
        selection_set_name = "SelectionSet_%d" % len(FreeCAD.ActiveDocument.Objects)

    if force_mode not in (None, "faces", "solids"):
        raise ValueError("force_mode must be 'faces', 'solids', or None")
    if volume_mode is not None and volume_mode not in (
        "fully_inside",
        "intersection",
        "coordinates",
    ):
        raise ValueError(
            "volume_mode must be 'fully_inside', 'intersection', 'coordinates', or None"
        )

    if force_mode is not None:
        mode = force_mode
    else:
        # Ask user for selection mode (faces or solids); allow cancel/close
        mode_dialog = QtWidgets.QMessageBox()
        mode_dialog.setWindowTitle("SelectionSet Type")
        mode_dialog.setText("What do you want to store in the SelectionSet?")
        faces_btn = mode_dialog.addButton("Faces", QtWidgets.QMessageBox.AcceptRole)
        solids_btn = mode_dialog.addButton("Solids", QtWidgets.QMessageBox.AcceptRole)
        cancel_btn = mode_dialog.addButton("Cancel", QtWidgets.QMessageBox.RejectRole)
        mode_dialog.setDefaultButton(faces_btn)
        try:
            # Allow Esc / window close to behave like Cancel when supported
            mode_dialog.setEscapeButton(cancel_btn)
        except Exception:
            pass
        mode_dialog.exec_()
        clicked = mode_dialog.clickedButton()
        if clicked is None or clicked == cancel_btn:
            # User cancelled or closed the dialog: abort Create/Update SelectionSet
            return
        if clicked == faces_btn:
            mode = "faces"
        elif clicked == solids_btn:
            mode = "solids"
        else:
            # Fallback: default to faces if something unexpected happens
            mode = "faces"

    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")
    if SELECTIONSET_VERBOSE_CALLBACK:
        msg = "[DEBUG] SelectionSet button callback triggered!"
        if report:
            report.append(msg)
        else:
            print(msg)
        raw_sel = FreeCADGui.Selection.getSelection()
        raw_sel_names = [obj.Name for obj in raw_sel]
        msg = f"[DEBUG] FreeCADGui.Selection.getSelection(): {raw_sel_names}"
        if report:
            report.append(msg)
        else:
            print(msg)
        sel_ex = FreeCADGui.Selection.getSelectionEx()
        sel_ex_names = [s.ObjectName for s in sel_ex]
        msg = f"[DEBUG] FreeCADGui.Selection.getSelectionEx(): {sel_ex_names}"
        if report:
            report.append(msg)
        else:
            print(msg)

    sel = FreeCADGui.Selection.getSelectionEx()
    # SelectionSets whose ElementList is all solids are used as solid filter source; don't treat them as "to update"
    selection_sets_to_update = [
        s.Object
        for s in sel
        if _is_selectionset(s.Object) and not _is_solid_filter_selectionset(s.Object)
    ]
    if selection_sets_to_update:
        sel = [s for s in sel if s.Object not in selection_sets_to_update]
    elements = []
    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")
    if SELECTIONSET_VERBOSE_CALLBACK:
        msg_header = "\n--- SelectionSet Button Pressed ---"
        msg_sel_count = f"Selection count: {len(sel)}"
        msg_sel_details = "Selected objects (in order):"
        sel_names = []
        for idx, s in enumerate(sel):
            sel_names.append(
                f"[{idx}] {s.ObjectName} (type: {type(s.Object).__name__})"
            )
        msg_sel_details += "\n  " + "\n  ".join(sel_names)
        msg_subel = "Selected subelements (per object):"
        for idx, s in enumerate(sel):
            msg_subel += f"\n  [{idx}] {s.ObjectName}: {s.SubElementNames}"
        if report:
            report.append(msg_header)
            report.append(msg_sel_count)
            report.append(msg_sel_details)
            report.append(msg_subel)
        else:
            print(msg_header)
            print(msg_sel_count)
            print(msg_sel_details)
            print(msg_subel)

    solid_filter_list = []
    main_obj = None
    selection_shape = None
    use_coordinates_mode = volume_mode == "coordinates"
    if use_coordinates_mode and len(sel) >= 1:
        # Coordinates mode: MainObject from first selected; point from CoordinatePoint or second object's center
        main_obj = sel[0].Object
        if len(sel) >= 2 and getattr(sel[1].Object, "Shape", None):
            selection_shape = sel[1].Object
            try:
                s = selection_shape.Shape
                if hasattr(s, "CenterOfMass"):
                    pt = s.CenterOfMass
                elif getattr(s, "BoundBox", None):
                    bb = s.BoundBox
                    pt = FreeCAD.Vector(
                        0.5 * (bb.XMin + bb.XMax),
                        0.5 * (bb.YMin + bb.YMax),
                        0.5 * (bb.ZMin + bb.ZMax),
                    )
                else:
                    pt = None
                if pt is not None:
                    elements = (
                        get_faces_by_point(main_obj, pt, tol=1e-6)
                        if mode == "faces"
                        else get_solids_by_point(main_obj, pt, tol=1e-6)
                    )
                    if SELECTIONSET_VERBOSE_CALLBACK and report:
                        report.append(
                            "Coordinates mode: point from shape '%s' center; %d elements."
                            % (selection_shape.Name, len(elements))
                        )
            except Exception as e:
                if SELECTIONSET_VERBOSE_CALLBACK and report:
                    report.append(
                        "Coordinates mode: could not get point from shape: %s" % e
                    )
        # else: elements stay [], user sets CoordinatePoint and Update
    elif mode == "faces":
        if len(sel) >= 3:
            # Find the SelectionSet whose ElementList is solids (solid filter); the other two are main and shape
            for s in sel:
                o = s.Object
                if hasattr(o, "ElementList"):
                    third_list = list(o.ElementList or [])
                    if third_list and all(".Solid" in str(e) for e in third_list):
                        solid_filter_list = third_list
                        if SELECTIONSET_VERBOSE_CALLBACK and report:
                            report.append(
                                "Using solid filter from SelectionSet '%s': %s"
                                % (o.Name, solid_filter_list)
                            )
                        break
            # The two objects that are not the solid-filter SelectionSet: one is main (many faces), one is shape
            others = [
                s.Object
                for s in sel
                if not (
                    hasattr(s.Object, "ElementList")
                    and (
                        list(s.Object.ElementList or [])
                        and all(
                            ".Solid" in str(e) for e in (s.Object.ElementList or [])
                        )
                    )
                )
            ]
            if len(others) >= 2:

                def _face_count(obj):
                    try:
                        return (
                            len(obj.Shape.Faces)
                            if (
                                getattr(obj, "Shape", None)
                                and getattr(obj.Shape, "Faces", None)
                            )
                            else 0
                        )
                    except Exception:
                        return 0

                # Main object is the one with more faces (the boolean/part); selection shape is the other (e.g. Body)
                others.sort(key=_face_count, reverse=True)
                main_obj = others[0]
                selection_shape = others[1]
        if main_obj is None and len(sel) >= 2:
            main_obj = sel[0].Object
            selection_shape = sel[1].Object
        if (
            main_obj is not None
            and selection_shape is not None
            and getattr(selection_shape, "Shape", None)
        ):
            face_volume_mode = volume_mode or "fully_inside"
            elements = get_faces_by_selection_shape(
                main_obj,
                selection_shape.Shape,
                tol=1e-7,
                volume_mode=face_volume_mode,
                solid_filter_list=solid_filter_list if solid_filter_list else None,
            )
            if SELECTIONSET_VERBOSE_CALLBACK:
                msg = f"Volume-based selection ({face_volume_mode}): main object '{main_obj.Name}', selection-defining shape '{selection_shape.Name}'."
                if solid_filter_list:
                    msg += f" Solid filter: {solid_filter_list}."
                msg += f"\nSelected faces: {elements}"
                print(msg)
                if report:
                    report.append(msg)
        else:
            for s in sel:
                objName = s.ObjectName
                for el in s.SubElementNames:
                    elements.append(f"{objName}.{el}")
            if SELECTIONSET_VERBOSE_CALLBACK:
                msg = f"Direct selection: elements {elements}"
                if report:
                    report.append(msg)
                else:
                    print(msg)
    else:  # mode == "solids"
        if len(sel) == 2:
            main_obj = sel[0].Object
            selection_shape = sel[1].Object
            solid_volume_mode = volume_mode or "fully_inside"
            # fully_inside: solid fully inside shape; intersection: solid has any overlap with shape
            for i, solid in enumerate(main_obj.Shape.Solids):
                if solid_volume_mode == "intersection":
                    include = _solid_intersects_shape(solid, selection_shape.Shape)
                else:
                    include = _solid_fully_inside_shape(solid, selection_shape.Shape)
                if SELECTIONSET_VERBOSE_CALLBACK:
                    msg = f"Solid {i + 1}: ({solid_volume_mode}) include={include}"
                    if report:
                        report.append(msg)
                    else:
                        print(msg)
                if include:
                    elements.append(f"{main_obj.Name}.Solid{i + 1}")
            if SELECTIONSET_VERBOSE_CALLBACK:
                msg = f"Volume-based solid selection ({solid_volume_mode}): main object '{main_obj.Name}', selection-defining shape '{selection_shape.Name}'.\nSelected solids: {elements}"
                if report:
                    report.append(msg)
                else:
                    print(msg)
        else:
            for s in sel:
                obj = s.Object
                if hasattr(obj, "Shape") and hasattr(obj.Shape, "Solids"):
                    for i, solid in enumerate(obj.Shape.Solids):
                        elements.append(f"{obj.Name}.Solid{i + 1}")
            if SELECTIONSET_VERBOSE_CALLBACK:
                msg = f"Direct solid selection: elements {elements}"
                if report:
                    report.append(msg)
                else:
                    print(msg)

    # If one or more SelectionSets were selected, update them from current selection instead of creating a new one
    if selection_sets_to_update:
        doc = FreeCAD.ActiveDocument
        for set_obj in selection_sets_to_update:
            if main_obj is not None and selection_shape is not None:
                set_obj.MainObject = main_obj
                set_obj.SelectionShape = selection_shape
                set_obj.Mode = mode
                set_obj.VolumeMode = volume_mode if volume_mode else "fully_inside"
                if use_coordinates_mode:
                    set_obj.VolumeMode = "coordinates"
                    if selection_shape and hasattr(selection_shape, "Shape"):
                        try:
                            s = selection_shape.Shape
                            if hasattr(s, "CenterOfMass"):
                                set_obj.CoordinatePoint = s.CenterOfMass
                            elif getattr(s, "BoundBox", None):
                                bb = s.BoundBox
                                set_obj.CoordinatePoint = FreeCAD.Vector(
                                    0.5 * (bb.XMin + bb.XMax),
                                    0.5 * (bb.YMin + bb.YMax),
                                    0.5 * (bb.ZMin + bb.ZMax),
                                )
                        except Exception:
                            pass
                if solid_filter_list and hasattr(set_obj, "SolidFilterRefs"):
                    try:
                        refs = []
                        for s in solid_filter_list:
                            s = (s or "").strip()
                            if ".Solid" in s and "." in s:
                                name, sub = s.split(".", 1)
                                o = doc.getObject(name) if doc else None
                                if o and sub:
                                    refs.append((o, sub))
                        if refs:
                            set_obj.SolidFilterRefs = refs
                    except Exception:
                        pass
            recomputeSelectionSetFromShapes(set_obj)
            apply_selectionset_view_defaults(set_obj)
        if SELECTIONSET_VERBOSE_CALLBACK:
            msg = "Updated %d SelectionSet(s) from current selection." % len(
                selection_sets_to_update
            )
            if report:
                report.append(msg)
            else:
                print(msg)
        return selection_sets_to_update[0]

    obj = FreeCAD.ActiveDocument.addObject("App::FeaturePython", selection_set_name)
    SelectionSet(obj)
    if obj.ViewObject:
        obj.ViewObject.Visibility = (
            True  # avoid grayed-out in tree (FeaturePython defaults to hidden)
        )
        if hasattr(obj, "Visibility"):
            obj.Visibility = True
    # Ensure GUI document view is updated so tree does not stay grayed
    try:
        gui_doc = FreeCADGui.ActiveDocument
        if gui_doc and hasattr(gui_doc, "getViewObject"):
            vo = gui_doc.getViewObject(obj.Name)
            if vo is not None:
                vo.Visibility = True
                if hasattr(vo, "touch"):
                    vo.touch()
    except Exception:
        pass
    # Tree label: "SelectionSet ..." (suffix from name)
    _suffix = (
        selection_set_name[13:]
        if selection_set_name.startswith("SelectionSet_")
        else selection_set_name
    )
    obj.Label = "SelectionSet " + _suffix
    obj.ElementList = elements
    # Store the two shapes and mode when volume-based (two or three objects) or coordinates mode; optional solid filter when three
    if use_coordinates_mode and main_obj is not None:
        obj.MainObject = main_obj
        obj.Mode = mode
        obj.VolumeMode = "coordinates"
        if selection_shape is not None:
            obj.SelectionShape = selection_shape
            try:
                s = selection_shape.Shape
                if hasattr(s, "CenterOfMass"):
                    obj.CoordinatePoint = s.CenterOfMass
                elif getattr(s, "BoundBox", None):
                    bb = s.BoundBox
                    obj.CoordinatePoint = FreeCAD.Vector(
                        0.5 * (bb.XMin + bb.XMax),
                        0.5 * (bb.YMin + bb.YMax),
                        0.5 * (bb.ZMin + bb.ZMax),
                    )
            except Exception:
                pass
        # else: CoordinatePoint stays (0,0,0), user sets it and Update
    elif len(sel) >= 2:
        if main_obj is None:
            main_obj = sel[0].Object
            selection_shape = sel[1].Object
        if main_obj is not None and selection_shape is not None:
            obj.MainObject = main_obj
            obj.SelectionShape = selection_shape
            obj.Mode = mode
            obj.VolumeMode = volume_mode if volume_mode else "fully_inside"
            if solid_filter_list and hasattr(obj, "SolidFilterRefs"):
                try:
                    doc = FreeCAD.ActiveDocument
                    refs = []
                    for s in solid_filter_list:
                        s = (s or "").strip()
                        if ".Solid" in s and "." in s:
                            name, sub = s.split(".", 1)
                            o = doc.getObject(name) if doc else None
                            if o and sub:
                                refs.append((o, sub))
                    if refs:
                        obj.SolidFilterRefs = refs
                except Exception:
                    pass
    FreeCAD.ActiveDocument.recompute()
    if SELECTIONSET_VERBOSE_CALLBACK:
        msg = f"SelectionSet created with {len(elements)} elements."
        if report:
            report.append(msg)
        else:
            print(msg)
    return obj


def recomputeSelectionSetFromShapes(set_obj):
    """
    Recompute ElementList from the stored MainObject and SelectionShape (and Mode, VolumeMode).
    Use this after changing MainObject or SelectionShape in the property panel.
    For VolumeMode "coordinates", only MainObject is required; point comes from CoordinatePoint or SelectionShape center.
    No-op if MainObject is not set (or for volume modes, if SelectionShape is not set).
    """
    if not hasattr(set_obj, "MainObject"):
        return False
    main_obj = set_obj.MainObject
    if not main_obj:
        return False
    mode = getattr(set_obj, "Mode", None) or "faces"
    volume_mode = getattr(set_obj, "VolumeMode", None) or "fully_inside"
    if volume_mode not in ("fully_inside", "intersection", "coordinates"):
        volume_mode = "fully_inside"

    if volume_mode == "coordinates":
        point = _get_coordinate_point_from_set(set_obj)
        if point is None:
            report = FreeCADGui.getMainWindow().findChild(
                QtWidgets.QTextEdit, "Report view"
            )
            msg = "[SelectionSet] Coordinates mode: set CoordinatePoint (x,y,z) or set SelectionShape to use its center."
            print(msg)
            if report:
                report.append(msg)
            set_obj.ElementList = []
            FreeCAD.ActiveDocument.recompute()
            return True
        solid_filter_list = _get_solid_filter_list_from_set(set_obj)
        if mode == "faces":
            elements = get_faces_by_point(
                main_obj,
                point,
                tol=1e-6,
                solid_filter_list=solid_filter_list if solid_filter_list else None,
            )
        else:
            elements = get_solids_by_point(main_obj, point, tol=1e-6)
        set_obj.ElementList = elements
        FreeCAD.ActiveDocument.recompute()
        return True

    # Volume-based mode: need SelectionShape
    if not hasattr(set_obj, "SelectionShape"):
        return False
    selection_shape = set_obj.SelectionShape
    if not selection_shape:
        return False
    if volume_mode not in ("fully_inside", "intersection"):
        volume_mode = "fully_inside"
    elements = []
    solid_filter_list = _get_solid_filter_list_from_set(set_obj)
    if mode == "faces":
        elements = get_faces_by_selection_shape(
            main_obj,
            selection_shape.Shape,
            tol=1e-7,
            volume_mode=volume_mode,
            solid_filter_list=solid_filter_list if solid_filter_list else None,
        )
    else:
        solid_volume_mode = (
            volume_mode
            if volume_mode in ("fully_inside", "intersection")
            else "fully_inside"
        )
        for i, solid in enumerate(main_obj.Shape.Solids):
            if solid_volume_mode == "intersection":
                include = _solid_intersects_shape(solid, selection_shape.Shape)
            else:
                include = _solid_fully_inside_shape(solid, selection_shape.Shape)
            if include:
                elements.append(f"{main_obj.Name}.Solid{i + 1}")
    set_obj.ElementList = elements
    FreeCAD.ActiveDocument.recompute()
    return True


def create_selectionset(
    doc=None,
    name=None,
    main_object=None,
    selection_shape=None,
    coordinate_point=None,
    mode="faces",
    volume_mode="fully_inside",
    solid_filter_refs=None,
    label=None,
):
    """
    Create a SelectionSet from Python (no GUI selection required).

    doc: FreeCAD document (default ActiveDocument).
    name: object name (default "SelectionSet_1" etc.).
    main_object: Part object whose faces or solids are selected.
    selection_shape: for volume_mode "fully_inside" or "intersection", the shape that filters main_object. For "coordinates", optional: used as point source if coordinate_point is None.
    coordinate_point: for volume_mode "coordinates", the point (FreeCAD.Vector). If None and selection_shape has a shape, its center is used.
    mode: "faces" or "solids".
    volume_mode: "fully_inside", "intersection", or "coordinates".
    solid_filter_refs: optional list of (object, "Solid1") for face mode to restrict to those solids.
    label: tree label (default "SelectionSet " + name suffix).

    Returns the new SelectionSet object or None.
    """
    doc = doc or FreeCAD.ActiveDocument
    if not doc:
        FreeCAD.Console.PrintWarning("create_selectionset: no document.\n")
        return None
    if not main_object:
        FreeCAD.Console.PrintWarning("create_selectionset: main_object is required.\n")
        return None
    if name is None:
        name = "SelectionSet_%d" % (
            len(
                [
                    o
                    for o in doc.Objects
                    if getattr(o, "TypeId", "") == "App::FeaturePython"
                ]
            )
            + 1
        )
    existing = doc.getObject(name)
    if existing:
        name = name + "_%d" % len(doc.Objects)
    obj = doc.addObject("App::FeaturePython", name)
    SelectionSet(obj)
    if obj.ViewObject:
        obj.ViewObject.Visibility = True
    if hasattr(obj, "Visibility"):
        obj.Visibility = True
    obj.MainObject = main_object
    obj.Mode = mode
    obj.VolumeMode = volume_mode or "fully_inside"
    if volume_mode == "coordinates":
        if coordinate_point is not None:
            obj.CoordinatePoint = FreeCAD.Vector(
                coordinate_point.x, coordinate_point.y, coordinate_point.z
            )
        if selection_shape and getattr(selection_shape, "Shape", None):
            try:
                s = selection_shape.Shape
                if hasattr(s, "CenterOfMass"):
                    obj.CoordinatePoint = s.CenterOfMass
                elif getattr(s, "BoundBox", None):
                    bb = s.BoundBox
                    obj.CoordinatePoint = FreeCAD.Vector(
                        0.5 * (bb.XMin + bb.XMax),
                        0.5 * (bb.YMin + bb.YMax),
                        0.5 * (bb.ZMin + bb.ZMax),
                    )
            except Exception:
                pass
        if selection_shape:
            obj.SelectionShape = selection_shape
    else:
        if selection_shape:
            obj.SelectionShape = selection_shape
        if solid_filter_refs and hasattr(obj, "SolidFilterRefs"):
            obj.SolidFilterRefs = list(solid_filter_refs)
    _suffix = name[13:] if name.startswith("SelectionSet_") else name
    obj.Label = label if label is not None else ("SelectionSet " + _suffix)
    recomputeSelectionSetFromShapes(obj)
    apply_selectionset_view_defaults(obj)
    return obj


def selectElementsFromSet(set_obj):
    # Use the new showSelection method for best highlighting
    if hasattr(set_obj, "Proxy") and hasattr(set_obj.Proxy, "showSelection"):
        set_obj.Proxy.showSelection(set_obj)
    else:
        # Fallback: just select
        FreeCADGui.Selection.clearSelection()
        for entry in set_obj.ElementList:
            if "." in entry:
                obj_name, sub = entry.split(".", 1)
                obj = FreeCAD.ActiveDocument.getObject(obj_name)
                if obj:
                    FreeCADGui.Selection.addSelection(obj, sub)
            else:
                obj = FreeCAD.ActiveDocument.getObject(entry)
                if obj:
                    FreeCADGui.Selection.addSelection(obj)


def _add_elements_from_set_to_selection(set_obj, clear_first=False):
    """Add all elements in set_obj.ElementList to the current selection. If clear_first, clear selection first."""
    if clear_first:
        FreeCADGui.Selection.clearSelection()
    for entry in set_obj.ElementList or []:
        if "." in entry:
            obj_name, sub = entry.split(".", 1)
            obj = FreeCAD.ActiveDocument.getObject(obj_name)
            if obj:
                FreeCADGui.Selection.addSelection(obj, sub)
        else:
            obj = FreeCAD.ActiveDocument.getObject(entry)
            if obj:
                FreeCADGui.Selection.addSelection(obj)


def _add_refs_to_selection(refs, clear_first=False):
    """Add a list of (object, subelement_name) refs to the current GUI selection. If clear_first, clear selection first."""
    if clear_first:
        FreeCADGui.Selection.clearSelection()
    for obj, sub in refs or []:
        if obj is None:
            continue
        if sub:
            FreeCADGui.Selection.addSelection(obj, sub)
        else:
            FreeCADGui.Selection.addSelection(obj)


def _add_main_shape_to_selection(set_obj, clear_first=False):
    """Add MainObject of the SelectionSet to the current selection."""
    if clear_first:
        FreeCADGui.Selection.clearSelection()
    main = getattr(set_obj, "MainObject", None)
    if main:
        FreeCADGui.Selection.addSelection(main)


def _add_selection_shape_to_selection(set_obj, clear_first=False):
    """Add SelectionShape of the SelectionSet to the current selection."""
    if clear_first:
        FreeCADGui.Selection.clearSelection()
    sel_shape = getattr(set_obj, "SelectionShape", None)
    if sel_shape:
        FreeCADGui.Selection.addSelection(sel_shape)


def _add_both_shapes_to_selection(set_obj, clear_first=False):
    """Add MainObject and SelectionShape to the current selection."""
    if clear_first:
        FreeCADGui.Selection.clearSelection()
    main = getattr(set_obj, "MainObject", None)
    sel_shape = getattr(set_obj, "SelectionShape", None)
    if main:
        FreeCADGui.Selection.addSelection(main)
    if sel_shape:
        FreeCADGui.Selection.addSelection(sel_shape)


def _create_link_from_selectionset_menu(set_obj):
    """
    Create a SelectionSetLink with this SelectionSet as source (context menu).
    If a FEM object was selected, it is set as target and refs are applied.
    """
    link = createSelectionSetLinkFromSet(set_obj)
    if not link:
        FreeCAD.Console.PrintWarning(
            "Create SelectionSetLink: failed (no document or set).\n"
        )
        return
    if getattr(link, "TargetObjects", None) and len(link.TargetObjects) > 0:
        try:
            _apply_link_to_target(link)
            targets = link.Proxy.get_targets(link)
            names = ", ".join((t.Label or t.Name) for t in targets[:3])
            if len(targets) > 3:
                names += ", ..."
            FreeCAD.Console.PrintMessage(
                "SelectionSetLink '%s' created with this set and target(s) %s; refs applied.\n"
                % (link.Name, names)
            )
        except Exception as e:
            FreeCAD.Console.PrintWarning(
                "SelectionSetLink created; apply to target failed: %s\n" % e
            )
    else:
        FreeCAD.Console.PrintMessage(
            "SelectionSetLink '%s' created with this set. Select FEM object(s) (e.g. constraint, material) and add them to Target objects in the Data tab (click ...), or multi-select set + FEM and use this menu again.\n"
            % link.Name
        )


def _update_selectionset_from_menu(set_obj):
    """Recompute this SelectionSet's ElementList from its stored MainObject and SelectionShape (context menu)."""
    if recomputeSelectionSetFromShapes(set_obj):
        FreeCAD.Console.PrintMessage(
            "SelectionSet '%s' updated from stored shapes.\n"
            % (set_obj.Name or "SelectionSet")
        )
    else:
        FreeCAD.Console.PrintWarning(
            "SelectionSet '%s': could not update (MainObject or SelectionShape missing).\n"
            % (set_obj.Name or "SelectionSet")
        )


def _delete_selectionset_only(set_obj):
    """
    Delete only the SelectionSet object itself, keeping all shapes and links.
    This calls removeObject on the SelectionSet by name, so FreeCAD does not
    try to delete its MainObject/SelectionShape first.
    """
    doc = getattr(set_obj, "Document", None) or FreeCAD.ActiveDocument
    if not doc:
        return
    name = set_obj.Name
    _log_to_report(
        "[SelectionSet] Delete-only requested for '%s' (keep shapes and links)." % name
    )
    try:
        doc.removeObject(name)
        doc.recompute()
    except Exception as e:
        FreeCAD.Console.PrintWarning(
            "[SelectionSet] Delete-only failed for '%s': %s\n" % (name, e)
        )


# ---------- FEM integration: use SelectionSet as references for constraints/loads ----------
def get_fem_references_from_selectionset(set_obj):
    """
    Return FEM-style References from a SelectionSet: list of (object, subelement_name) tuples.
    FEM constraints expect References as [(obj, "Face1"), (obj, "Edge2"), ...].
    Returns [] if set_obj has no ElementList or entries cannot be resolved.
    """
    refs = []
    for entry in getattr(set_obj, "ElementList", None) or []:
        if "." in entry:
            obj_name, sub_name = entry.split(".", 1)
            obj = FreeCAD.ActiveDocument.getObject(obj_name)
            if obj:
                refs.append((obj, sub_name))
        else:
            obj = FreeCAD.ActiveDocument.getObject(entry)
            if obj:
                refs.append((obj, ""))
    return refs


def get_fem_references_from_selectionsets(set_objs):
    """
    Return combined FEM-style References from multiple SelectionSets.
    set_objs: list of SelectionSet objects (or single object).
    Order: refs from first set, then second, etc. No deduplication.
    """
    if not hasattr(set_objs, "__iter__") or isinstance(set_objs, (str, bytes)):
        set_objs = [set_objs]
    refs = []
    for set_obj in set_objs:
        refs.extend(get_fem_references_from_selectionset(set_obj))
    return refs


def apply_selectionset_to_fem_constraint(set_obj, constraint):
    """
    Set a FEM constraint's References from a SelectionSet's ElementList.
    constraint: any document object with a References property (e.g. Fem::ConstraintFixed).
    Returns True if references were set, False if constraint has no References or refs are empty.
    """
    refs = get_fem_references_from_selectionset(set_obj)
    if not refs:
        return False
    if not hasattr(constraint, "References"):
        return False
    constraint.References = refs
    constraint.touch()
    if getattr(constraint, "Document", None):
        constraint.Document.recompute()
    return True


def apply_references_to_fem_constraint_one_by_one(refs, constraint, doc=None):
    """
    Set a FEM constraint's References by appending one (object, subelement) at a time.

    Some FEM constraints (e.g. Fem::ConstraintElectrostaticPotential) only accept or
    persist references when they are added one by one (GUI or script). Assigning a full
    list at once may result in only the first face being stored. This workaround builds
    References by repeatedly appending a single reference and touching the constraint,
    so all faces are retained.

    refs: list of (object, subelement_name) tuples.
    constraint: document object with References property.
    doc: optional document to recompute (default: constraint.Document).
    Returns True if constraint has References and refs were applied.
    """
    if not hasattr(constraint, "References"):
        return False
    current = list(getattr(constraint, "References", []) or [])
    if _refs_equal(current, refs):
        return True
    doc = doc or getattr(constraint, "Document", None)
    constraint.References = []
    constraint.touch()
    for obj, sub in refs:
        current = list(getattr(constraint, "References", []) or [])
        constraint.References = current + [(obj, sub)]
        constraint.touch()
        # Do not recompute here: when called from SelectionSetLink.execute() this would cause
        # recursive document recompute and freeze the GUI. Caller may call doc.recompute() once.
    return True


# ---------- SelectionSetLink: combine add/subtract SelectionSets and apply to target (e.g. FEM) ----------
def _refs_equal(refs_a, refs_b):
    """Return True if two FEM ref lists are equal (same objects and subnames). Compare by (obj.Name, sub)."""
    if refs_a is refs_b:
        return True
    if len(refs_a) != len(refs_b):
        return False
    norm = lambda r: (
        r[0].Name if hasattr(r[0], "Name") else r[0],
        r[1] if len(r) > 1 else "",
    )
    return [norm(r) for r in refs_a] == [norm(r) for r in refs_b]


def _is_selectionset(obj):
    """Return True if obj looks like a SelectionSet (has ElementList)."""
    return obj is not None and hasattr(obj, "ElementList")


def _format_refs_summary(refs):
    """Return a short summary string for a list of (obj, sub) refs, e.g. '6 faces' or '2 solids, 4 faces'."""
    if not refs:
        return "0 elements"
    counts = {}
    for _obj, sub in refs:
        if sub and "." not in sub:
            if sub.startswith("Face"):
                counts["face"] = counts.get("face", 0) + 1
            elif sub.startswith("Solid"):
                counts["solid"] = counts.get("solid", 0) + 1
            elif sub.startswith("Edge"):
                counts["edge"] = counts.get("edge", 0) + 1
            elif sub.startswith("Vertex"):
                counts["vertex"] = counts.get("vertex", 0) + 1
            else:
                counts["element"] = counts.get("element", 0) + 1
        else:
            counts["element"] = counts.get("element", 0) + 1
    parts = []
    if counts.get("face"):
        parts.append("%d face%s" % (counts["face"], "s" if counts["face"] != 1 else ""))
    if counts.get("solid"):
        parts.append(
            "%d solid%s" % (counts["solid"], "s" if counts["solid"] != 1 else "")
        )
    if counts.get("edge"):
        parts.append("%d edge%s" % (counts["edge"], "s" if counts["edge"] != 1 else ""))
    if counts.get("vertex"):
        parts.append(
            "%d vertex" % counts["vertex"] + ("es" if counts["vertex"] != 1 else "")
        )
    if counts.get("element"):
        parts.append(
            "%d element%s" % (counts["element"], "s" if counts["element"] != 1 else "")
        )
    return ", ".join(parts) if parts else "%d refs" % len(refs)


def _update_link_ref_summary(link_obj, refs):
    """Set link_obj.RefSummary and append hint to Label so the tree shows type/count (e.g. '6 faces')."""
    summary = _format_refs_summary(refs)
    try:
        link_obj.RefSummary = summary
    except Exception:
        pass
    try:
        label = (
            getattr(link_obj, "Label", None)
            or getattr(link_obj, "Name", "")
            or "SelectionSetLink"
        )
        # Strip previous parenthesized hint so we don't stack " (6 faces) (6 faces)"
        if " (" in label and label.rstrip().endswith(")"):
            label = label[: label.rfind(" (")].strip() or link_obj.Name
        link_obj.Label = (label + " (" + summary + ")").strip()
    except Exception:
        pass


def _get_refs_from_shapes(shapes_list, mode="faces"):
    """
    Convert a list of shape objects (Part, etc.) to FEM-style refs [(obj, subname), ...].
    mode: "faces" -> all faces (Obj.Face1, ...); "solids" -> all solids (Obj.Solid1, ...).
    Objects without Shape are skipped.
    """
    refs = []
    for o in shapes_list or []:
        if o is None or not getattr(o, "Shape", None):
            continue
        try:
            shape = o.Shape
            if mode == "solids" and getattr(shape, "Solids", None):
                for i in range(len(shape.Solids)):
                    refs.append((o, "Solid%d" % (i + 1)))
            elif getattr(shape, "Faces", None):
                for i in range(len(shape.Faces)):
                    refs.append((o, "Face%d" % (i + 1)))
        except Exception:
            continue
    return refs


def _target_and_add_sets_from_selection():
    """
    From the current GUI selection, pick target(s) (objects with References),
    add-list (SelectionSets), and add-shapes (objects with Shape that are not SelectionSets).
    Returns (first_target or None, list_of_selection_sets, list_of_shapes, list_of_all_targets).
    """
    try:
        sel = FreeCADGui.Selection.getSelection()
    except Exception:
        return (None, [], [], [])
    targets = []
    add_sets = []
    add_shapes = []
    for obj in sel:
        if obj is None:
            continue
        if hasattr(obj, "References"):
            targets.append(obj)
        if _is_selectionset(obj):
            add_sets.append(obj)
        elif getattr(obj, "Shape", None) is not None:
            add_shapes.append(obj)
    first = targets[0] if targets else None
    return (first, add_sets, add_shapes, targets)


class SelectionSetLink:
    """
    Tree object that combines elements from SelectionSets and/or direct shapes (add lists),
    subtracts elements from SelectionSets and/or direct shapes (subtract lists), and applies
    the result to a target object (e.g. FEM constraint). Both selection sets and shapes
    are evaluated together. Right-click the link → SelectionSet link → Renew elements.

    Refresh behaviour: The target (FEM constraint, etc.) is updated only when the user
    chooses "Renew elements" (right-click link → SelectionSet link). We do not apply
    on property change or document recompute, to avoid "still touched after recompute"
    and repeated recomputes on FEM objects. After changing the link's inputs or linked
    set content, use "Renew elements" to push refs to the target.
    """

    # Properties that, when changed, should re-apply the combined refs to the target.
    _INPUT_PROPS = (
        "TargetObjects",
        "AddSelectionSets",
        "AddShapes",
        "SubtractSelectionSets",
        "SubtractShapes",
        "ShapeRefMode",
    )

    def __init__(self, obj):
        # Order: targets first, then add lists, then subtract lists (Data tab sub-view)
        obj.addProperty(
            "App::PropertyLinkList",
            "TargetObjects",
            "SelectionSet link",
            "Targets: FEM elements (constraints, material, etc.) whose References will be set. Click ... to add one or more; same combined refs are applied to each. Right-click link → Renew elements to apply.",
        )
        obj.addProperty(
            "App::PropertyLinkList",
            "AddSelectionSets",
            "SelectionSet link",
            "Selection sets to add: list of SelectionSet objects whose elements are combined (order preserved)",
        )
        obj.addProperty(
            "App::PropertyLinkList",
            "AddShapes",
            "SelectionSet link",
            "Shapes to add: Part/geometry objects; all faces or solids (see Shape ref mode) are added",
        )
        obj.addProperty(
            "App::PropertyLinkList",
            "SubtractSelectionSets",
            "SelectionSet link",
            "Selection sets to subtract: elements in these sets are removed from the result",
        )
        obj.addProperty(
            "App::PropertyLinkList",
            "SubtractShapes",
            "SelectionSet link",
            "Shapes to subtract: Part/geometry objects; their faces or solids are removed from the result",
        )
        obj.addProperty(
            "App::PropertyEnumeration",
            "ShapeRefMode",
            "SelectionSet link",
            "For direct shapes: add/subtract all Faces or all Solids",
        )
        obj.ShapeRefMode = ["faces", "solids"]
        obj.ShapeRefMode = "faces"
        obj.addProperty(
            "App::PropertyBool",
            "UseOneByOneWorkaround",
            "SelectionSet link",
            "If true, insert references one-by-one (for constraints that only accept one face at a time)",
        )
        obj.UseOneByOneWorkaround = True
        obj.addProperty(
            "App::PropertyString",
            "RefSummary",
            "SelectionSet link",
            "Summary of combined refs (e.g. '6 faces', '2 solids, 4 faces'); updated when refs are computed",
        )
        obj.setEditorMode("RefSummary", 1)  # ReadOnly
        # Meta: macro version when this object was created (read-only)
        try:
            obj.addProperty(
                "App::PropertyString",
                "MacroVersion",
                "Meta",
                "SelectionSet macro version that created this object.",
            )
            obj.MacroVersion = __version__
            obj.setEditorMode("MacroVersion", 1)  # ReadOnly
        except Exception:
            pass
        obj.Proxy = self

    def onDocumentRestored(self, obj):
        if not hasattr(obj, "TargetObjects"):
            obj.addProperty(
                "App::PropertyLinkList",
                "TargetObjects",
                "SelectionSet link",
                "Targets: FEM elements (constraints, material, etc.) whose References will be set. Same combined refs are applied to each.",
            )
        if not hasattr(obj, "AddShapes"):
            obj.addProperty(
                "App::PropertyLinkList",
                "AddShapes",
                "SelectionSet link",
                "Shapes to add: Part/geometry objects; all faces or solids (see Shape ref mode) are added",
            )
        if not hasattr(obj, "SubtractShapes"):
            obj.addProperty(
                "App::PropertyLinkList",
                "SubtractShapes",
                "SelectionSet link",
                "Shapes to subtract: Part/geometry objects; their faces or solids are removed from the result",
            )
        if not hasattr(obj, "ShapeRefMode"):
            obj.addProperty(
                "App::PropertyEnumeration",
                "ShapeRefMode",
                "SelectionSet link",
                "For direct shapes: add/subtract all Faces or all Solids",
            )
            obj.ShapeRefMode = ["faces", "solids"]
            obj.ShapeRefMode = "faces"
        if not hasattr(obj, "UseOneByOneWorkaround"):
            obj.addProperty(
                "App::PropertyBool",
                "UseOneByOneWorkaround",
                "SelectionSet link",
                "If true, add references one-by-one (for constraints that only accept one face at a time)",
            )
        if not hasattr(obj, "RefSummary"):
            obj.addProperty(
                "App::PropertyString",
                "RefSummary",
                "SelectionSet link",
                "Summary of combined refs (e.g. '6 faces', '2 solids, 4 faces'); updated when refs are computed",
            )
            obj.setEditorMode("RefSummary", 1)
        if not hasattr(obj, "MacroVersion"):
            try:
                obj.addProperty(
                    "App::PropertyString",
                    "MacroVersion",
                    "Meta",
                    "SelectionSet macro version that created this object.",
                )
                obj.MacroVersion = __version__
                obj.setEditorMode("MacroVersion", 1)
            except Exception:
                pass
        obj.Proxy = self
        if (
            hasattr(obj, "ViewObject")
            and obj.ViewObject
            and not isinstance(
                getattr(obj.ViewObject, "Proxy", None), ViewProviderSelectionSetLink
            )
        ):
            obj.ViewObject.Proxy = ViewProviderSelectionSetLink(obj.ViewObject)

    def onChanged(self, obj, prop):
        """Update ref summary/label when link input properties change; do not apply to target (avoids FEM recompute issues)."""
        if prop in self._INPUT_PROPS:
            try:
                self.get_combined_refs(obj)
            except Exception:
                pass

    def execute(self, obj):
        """Do not apply to target here: FEM elements are only updated when the user chooses 'Renew elements'."""
        pass

    def get_combined_refs(self, obj):
        """
        Return list of (document_object, subelement_name) from AddSelectionSets + AddShapes
        minus (SubtractSelectionSets + SubtractShapes). Order: add sets first, then add shapes.
        """
        doc = obj.Document if hasattr(obj, "Document") else FreeCAD.ActiveDocument
        if not doc:
            return []
        mode = getattr(obj, "ShapeRefMode", "faces") or "faces"
        add_sets = [
            o
            for o in (getattr(obj, "AddSelectionSets", None) or [])
            if _is_selectionset(o)
        ]
        add_shapes = [
            o for o in (getattr(obj, "AddShapes", None) or []) if o is not None
        ]
        sub_sets = [
            o
            for o in (getattr(obj, "SubtractSelectionSets", None) or [])
            if _is_selectionset(o)
        ]
        sub_shapes = [
            o for o in (getattr(obj, "SubtractShapes", None) or []) if o is not None
        ]
        refs_add = get_fem_references_from_selectionsets(
            add_sets
        ) + _get_refs_from_shapes(add_shapes, mode)
        refs_sub = get_fem_references_from_selectionsets(
            sub_sets
        ) + _get_refs_from_shapes(sub_shapes, mode)
        sub_set = set((r[0].Name, r[1]) for r in refs_sub)
        result = []
        for r in refs_add:
            key = (r[0].Name, r[1])
            if key not in sub_set:
                result.append(r)
        try:
            _update_link_ref_summary(obj, result)
        except Exception:
            pass
        return result

    def get_targets(self, obj):
        """Return list of target objects (with References) from TargetObjects, deduplicated."""
        raw = getattr(obj, "TargetObjects", None) or []
        seen = set()
        out = []
        for t in raw:
            if t is None or id(t) in seen or not hasattr(t, "References"):
                continue
            seen.add(id(t))
            out.append(t)
        return out

    def apply_to_target(self, obj):
        """
        Apply combined refs to each target in TargetObjects. Returns True if applied to at least one.
        When combined refs are 0, each target's References are cleared (set to []).
        Skips a target when it already has the same refs, to avoid repeated "still touched after recompute".
        """
        targets = self.get_targets(obj)
        if not targets:
            return False
        refs = self.get_combined_refs(obj)
        doc = getattr(obj, "Document", None) or FreeCAD.ActiveDocument
        use_one_by_one = getattr(obj, "UseOneByOneWorkaround", True)
        applied = 0
        for target in targets:
            current = list(getattr(target, "References", []) or [])
            if _refs_equal(current, refs):
                applied += 1
                continue
            if use_one_by_one:
                if apply_references_to_fem_constraint_one_by_one(refs, target, doc):
                    applied += 1
            else:
                target.References = refs if refs else []
                target.touch()
                applied += 1
        # Do not call doc.recompute() here: callers (e.g. "Renew elements") must call it themselves.
        return applied > 0


# Icons: SelectionSetLink.svg references FEM_Analysis.svg by filename; if that file is in the same icons/ folder it is shown (no script, no download).


class ViewProviderSelectionSetLink:
    def __init__(self, vobj):
        vobj.Proxy = self
        # Set immediately so tree does not show grayed (App::FeaturePython defaults to Visibility=False)
        vobj.Visibility = True

    def getDisplayModes(self, vobj):
        """Return list of display mode names so DisplayMode property is non-empty (required for show/hide and tree)."""
        return ["Default"]

    def getIcon(self):
        """Return SelectionSetLink.svg, which links to FEM_Analysis.svg when that file is in the same folder."""
        icons_dir = _macro_icons_dir()
        for name in ("SelectionSetLink.svg", "selection_set_link.svg"):
            path = os.path.join(icons_dir, name)
            if path and os.path.isfile(path):
                return os.path.abspath(path)
        return ""

    def setupContextMenu(self, obj, menu):
        link_obj = obj.Object if hasattr(obj, "Object") else obj
        if not link_obj or not hasattr(link_obj, "Proxy"):
            return
        proxy = getattr(link_obj, "Proxy", None)
        if not proxy or not hasattr(proxy, "get_combined_refs"):
            return
        menu.addSeparator()
        submenu = QtWidgets.QMenu("SelectionSet link", menu)
        font = submenu.menuAction().font()
        font.setBold(True)
        submenu.menuAction().setFont(font)
        submenu.addAction("Renew elements", lambda: _apply_link_to_target(link_obj))
        submenu.addSeparator()
        submenu.addAction(
            "Expand to elements (replace selection)",
            lambda: _link_expand_to_selection(link_obj, clear_first=True),
        )
        submenu.addAction(
            "Add elements from list to selection",
            lambda: _link_expand_to_selection(link_obj, clear_first=False),
        )
        submenu.addAction(
            "Select only elements",
            lambda: _link_expand_to_selection(link_obj, clear_first=True),
        )
        menu.addMenu(submenu)

    def attach(self, vobj):
        self.ViewObject = vobj
        self.Object = vobj.Object
        try:
            from pivy import coin

            vobj.addDisplayMode(coin.SoSeparator(), "Default")
            if hasattr(vobj, "setDisplayMode"):
                vobj.setDisplayMode("Default")
            if hasattr(vobj, "DisplayMode"):
                vobj.DisplayMode = "Default"
        except Exception:
            pass
        # Set visibility last so the tree view shows the object as visible (not grayed).
        vobj.Visibility = True
        if (
            hasattr(vobj, "Object")
            and vobj.Object
            and hasattr(vobj.Object, "Visibility")
        ):
            try:
                vobj.Object.Visibility = True
            except Exception:
                pass

        def _set_display_mode_default():
            try:
                if (
                    hasattr(vobj, "DisplayMode")
                    and getattr(vobj, "DisplayMode", "") != "Default"
                ):
                    vobj.DisplayMode = "Default"
            except Exception:
                pass

        try:
            QtCore.QTimer.singleShot(100, _set_display_mode_default)
        except Exception:
            pass

    def __getstate__(self):
        return None

    def __setstate__(self, state):
        pass


def _apply_link_to_target(link_obj):
    """Apply SelectionSetLink's combined refs to its target(s) (Target objects)."""
    if not hasattr(link_obj, "Proxy") or not hasattr(link_obj.Proxy, "apply_to_target"):
        FreeCAD.Console.PrintWarning("SelectionSetLink: invalid object.\n")
        return
    if link_obj.Proxy.apply_to_target(link_obj):
        n_refs = len(link_obj.Proxy.get_combined_refs(link_obj))
        n_tgt = len(link_obj.Proxy.get_targets(link_obj))
        if n_tgt <= 1:
            FreeCAD.Console.PrintMessage(
                "SelectionSetLink '%s': applied %d references to target.\n"
                % (link_obj.Label or link_obj.Name, n_refs)
            )
        else:
            FreeCAD.Console.PrintMessage(
                "SelectionSetLink '%s': applied %d references to %d targets.\n"
                % (link_obj.Label or link_obj.Name, n_refs, n_tgt)
            )
        doc = getattr(link_obj, "Document", None)
        if doc:
            doc.recompute()
    else:
        FreeCAD.Console.PrintWarning(
            "SelectionSetLink '%s': no targets. Add one or more objects (e.g. FEM constraint, material) to Target objects in the Data tab, then right-click → Renew elements.\n"
            % (link_obj.Label or link_obj.Name)
        )


def _link_expand_to_selection(link_obj, clear_first=False):
    """Add SelectionSetLink's combined refs to the GUI selection (replace or add)."""
    if not hasattr(link_obj, "Proxy") or not hasattr(
        link_obj.Proxy, "get_combined_refs"
    ):
        return
    refs = link_obj.Proxy.get_combined_refs(link_obj)
    _add_refs_to_selection(refs, clear_first=clear_first)


def createSelectionSetLinkFromSet(selection_set_obj, name=None):
    """
    Create a SelectionSetLink with the given SelectionSet in AddSelectionSets.
    If FEM object(s) (with References) are in the current selection, they are set as Target objects.
    Used from the SelectionSet context menu (right-click → Create SelectionSetLink).
    Returns the new link object or None.
    """
    doc = FreeCAD.ActiveDocument
    if not doc or not selection_set_obj:
        return None
    target, add_sets, add_shapes, all_targets = _target_and_add_sets_from_selection()
    # This set is the one we right-clicked; ensure it is in add_sets (first), merge with any others from selection
    if selection_set_obj not in add_sets:
        add_sets = [selection_set_obj] + [o for o in add_sets if o != selection_set_obj]
    else:
        add_sets = [o for o in add_sets if o != selection_set_obj]
        add_sets.insert(0, selection_set_obj)
    return createSelectionSetLink(
        name=name,
        _target=target,
        _add_sets=add_sets,
        _add_shapes=add_shapes,
        _targets=all_targets,
    )


def createSelectionSetLink(
    name=None, _target=None, _add_sets=None, _add_shapes=None, _targets=None
):
    """
    Create a SelectionSetLink in the active document. If FEM object(s) (with References),
    SelectionSets, and/or shapes are currently selected,     they are set as Target objects, Add selection sets, and Add shapes. Returns the new object or None.
    Internal use: pass _target, _add_sets, _add_shapes, _targets to override selection-based values.
    """
    doc = FreeCAD.ActiveDocument
    if not doc:
        FreeCAD.Console.PrintWarning("SelectionSetLink: no active document.\n")
        return None
    if (
        _target is not None
        or _add_sets is not None
        or _add_shapes is not None
        or _targets is not None
    ):
        target = _target
        add_sets = _add_sets if _add_sets is not None else []
        add_shapes = _add_shapes if _add_shapes is not None else []
        all_targets = (
            list(_targets) if _targets else ([target] if target is not None else [])
        )
    else:
        target, add_sets, add_shapes, all_targets = (
            _target_and_add_sets_from_selection()
        )
    if name is None:
        name = "SelectionSetLink_%d" % (len(doc.Objects) + 1)
    existing = doc.getObject(name)
    if existing:
        name = name + "_%d" % len(doc.Objects)
    obj = doc.addObject("App::FeaturePython", name)
    SelectionSetLink(obj)
    if obj.ViewObject:
        obj.ViewObject.Visibility = (
            True  # avoid grayed-out in tree (FeaturePython defaults to hidden)
        )
        if hasattr(obj, "Visibility"):
            obj.Visibility = True
    try:
        gui_doc = FreeCADGui.ActiveDocument
        if gui_doc and hasattr(gui_doc, "getViewObject"):
            vo = gui_doc.getViewObject(obj.Name)
            if vo is not None:
                vo.Visibility = True
                if hasattr(vo, "touch"):
                    vo.touch()
    except Exception:
        pass
    # Tree label: "SelectionSetLink ..." (suffix from name; ref summary will append " (N faces)" etc. later)
    _suffix = name[17:] if name.startswith("SelectionSetLink_") else name
    obj.Label = "SelectionSetLink " + _suffix
    obj.ViewObject.Proxy = ViewProviderSelectionSetLink(obj.ViewObject)
    if all_targets:
        obj.TargetObjects = list(all_targets)
    if add_sets:
        obj.AddSelectionSets = add_sets
    if add_shapes:
        obj.AddShapes = add_shapes
    doc.recompute()
    return obj


class CreateSelectionSetLinkCommand:
    """Command to create a SelectionSetLink in the tree."""

    def GetResources(self):
        return {
            "Pixmap": "",
            "MenuText": "Create SelectionSet link",
            "ToolTip": "Create a link object: add/subtract SelectionSets and apply result to a FEM or other target (References).",
        }

    def IsActive(self):
        return FreeCAD.ActiveDocument is not None

    def Activated(self):
        link = createSelectionSetLink()
        if link:
            has_target = bool(
                getattr(link, "TargetObjects", None) and len(link.TargetObjects) > 0
            )
            has_add = bool(
                getattr(link, "AddSelectionSets", None)
                and len(link.AddSelectionSets) > 0
            ) or bool(getattr(link, "AddShapes", None) and len(link.AddShapes) > 0)
            if has_target or has_add:
                FreeCAD.Console.PrintMessage(
                    "SelectionSetLink '%s' created from selection (target and/or add sets/shapes set). Right-click link → SelectionSet link → Renew elements to apply.\n"
                    % link.Name
                )
            else:
                FreeCAD.Console.PrintMessage(
                    "SelectionSetLink '%s' created. Add target(s) to Target objects (Data tab, click ...), set Add selection sets / Add shapes, then right-click link → Renew elements.\n"
                    % link.Name
                )


def update_all_selectionsets_and_links():
    """
    Update all SelectionSets in the active document (recompute from shapes), then apply all
    SelectionSetLinks to their targets. Run after geometry or selection changes.
    """
    doc = FreeCAD.ActiveDocument
    if not doc:
        FreeCAD.Console.PrintWarning("Update all: no active document.\n")
        return
    n_sets = 0
    for obj in doc.Objects:
        if getattr(obj, "Proxy", None) is None:
            continue
        if hasattr(obj, "ElementList") and hasattr(obj, "MainObject"):
            try:
                if recomputeSelectionSetFromShapes(obj):
                    n_sets += 1
            except Exception as e:
                FreeCAD.Console.PrintWarning(
                    "Update all: failed to update SelectionSet '%s': %s\n"
                    % (obj.Name, e)
                )
    if n_sets:
        doc.recompute()
    n_links = 0
    for obj in doc.Objects:
        if not hasattr(obj, "Proxy") or not hasattr(
            getattr(obj, "Proxy", None), "apply_to_target"
        ):
            continue
        try:
            _apply_link_to_target(obj)
            n_links += 1
        except Exception as e:
            FreeCAD.Console.PrintWarning(
                "Update all: failed to apply link '%s': %s\n" % (obj.Name, e)
            )
    if n_links:
        doc.recompute()
    FreeCAD.Console.PrintMessage(
        "Update all: %d SelectionSet(s) updated, %d SelectionSetLink(s) applied.\n"
        % (n_sets, n_links)
    )


# ---------- Toolbar: Create / Expand / Update SelectionSet buttons ----------
def add_selectionset_button_to_toolbar():
    """
    Adds a button to a FreeCAD toolbar called 'Selection_tools' to create a SelectionSet from the current selection.
    """
    mw = FreeCADGui.getMainWindow()
    toolbar_name = "Selection_tools"
    # Try to find the toolbar, or create it if it doesn't exist
    for child in mw.findChildren(QtWidgets.QToolBar):
        if child.objectName() == toolbar_name:
            toolbar = child
            break
    else:
        toolbar = QtWidgets.QToolBar(toolbar_name, mw)
        toolbar.setObjectName(toolbar_name)
        mw.addToolBar(toolbar)
    # Remove the old button if it exists (to allow reload)
    button_name = "CreateSelectionSetButton"
    old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
    if old_btn:
        # Remove the action associated with the old button
        action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
        if action:
            toolbar.removeAction(action)
        old_btn.deleteLater()
    # Add the new button
    btn = QtWidgets.QToolButton(toolbar)
    btn.setObjectName(button_name)
    icon = get_toolbar_icon("SelectionSet")
    if icon is not None:
        btn.setIcon(icon)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    btn.setText("Create/Update SelectionSet")
    btn.setToolTip(
        "If only SelectionSet(s) are selected: recompute from stored Main/Shape. Otherwise: create a new SelectionSet or update selected sets from current selection."
    )
    btn.clicked.connect(lambda: create_or_update_selectionset_from_current())
    toolbar.addWidget(btn)
    # Add "Expand SelectionSet" button (works on current selection; fallback when double-click doesn't)
    # Remove old Expand button if it exists (to avoid duplicates on macro reload)
    expand_button_name = "ExpandSelectionSetButton"
    old_expand_btn = toolbar.findChild(QtWidgets.QToolButton, expand_button_name)
    if old_expand_btn:
        action = (
            old_expand_btn.defaultAction()
            if hasattr(old_expand_btn, "defaultAction")
            else None
        )
        if action:
            toolbar.removeAction(action)
        old_expand_btn.deleteLater()
    expand_btn = QtWidgets.QToolButton(toolbar)
    expand_btn.setObjectName(expand_button_name)
    icon_expand = get_toolbar_icon("SelectionSet")
    if icon_expand is not None:
        expand_btn.setIcon(icon_expand)
        expand_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    expand_btn.setText("Expand SelectionSet")
    expand_btn.setToolTip(
        "Replace selection with the elements stored in the selected SelectionSet(s). Select a SelectionSet in the tree first."
    )
    expand_btn.clicked.connect(lambda: FreeCADGui.runCommand("HighlightSelectionSet"))
    toolbar.addWidget(expand_btn)
    # Create SelectionSet link (add/subtract SelectionSets → target FEM or other References)
    try:
        FreeCADGui.addCommand("CreateSelectionSetLink", CreateSelectionSetLinkCommand())
    except Exception:
        pass
    link_button_name = "CreateSelectionSetLinkButton"
    old_link_btn = toolbar.findChild(QtWidgets.QToolButton, link_button_name)
    if old_link_btn:
        action = (
            old_link_btn.defaultAction()
            if hasattr(old_link_btn, "defaultAction")
            else None
        )
        if action:
            toolbar.removeAction(action)
        old_link_btn.deleteLater()
    link_btn = QtWidgets.QToolButton(toolbar)
    link_btn.setObjectName(link_button_name)
    icon_link = get_toolbar_icon("SelectionSetLink")
    if icon_link is not None:
        link_btn.setIcon(icon_link)
        link_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    link_btn.setText("Create SelectionSet link")
    link_btn.setToolTip(
        "Create a link: combine add/subtract SelectionSets and apply to one or more targets. In Data tab set Target objects (click ... to add), Add selection sets; right-click link → Renew elements."
    )
    link_btn.clicked.connect(lambda: FreeCADGui.runCommand("CreateSelectionSetLink"))
    toolbar.addWidget(link_btn)
    # Update all SelectionSets then all SelectionSetLinks
    update_all_button_name = "UpdateAllSelectionSetsAndLinksButton"
    old_update_btn = toolbar.findChild(QtWidgets.QToolButton, update_all_button_name)
    if old_update_btn:
        action = (
            old_update_btn.defaultAction()
            if hasattr(old_update_btn, "defaultAction")
            else None
        )
        if action:
            toolbar.removeAction(action)
        old_update_btn.deleteLater()
    update_btn = QtWidgets.QToolButton(toolbar)
    update_btn.setObjectName(update_all_button_name)
    icon_upd = get_toolbar_icon("SelectionSet")
    if icon_upd is not None:
        update_btn.setIcon(icon_upd)
    update_btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    update_btn.setText("Update all sets and links")
    update_btn.setToolTip(
        "Recompute all SelectionSets from their shapes, then apply all SelectionSetLinks to their targets (run after geometry changes)."
    )
    update_btn.clicked.connect(lambda: update_all_selectionsets_and_links())
    toolbar.addWidget(update_btn)
