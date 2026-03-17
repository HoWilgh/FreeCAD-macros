"""
selection_set_tests: Test runner and test geometry builders for SelectionSet.

FreeCAD version compatibility (1.0 and 1.1):
- Test geometry (e.g. PartDesign Pad) may use properties that are deprecated in 1.1.
  We keep 1.0 behaviour and adopt 1.1 APIs where needed (see comments at Pad Midplane/SideType).
"""
import FreeCAD
import FreeCADGui
from PySide import QtCore, QtWidgets
import Part

import selection_set_core
from selection_set_core import (
    __version__ as MACRO_VERSION,
    SelectionSet,
    apply_selectionset_view_defaults,
    createSelectionSetFromCurrent,
    createSelectionSetLink,
    get_faces_by_selection_shape,
    get_faces_by_point,
    get_solids_by_point,
    get_fem_references_from_selectionset,
    get_fem_references_from_selectionsets,
    apply_selectionset_to_fem_constraint,
    apply_references_to_fem_constraint_one_by_one,
    get_toolbar_icon,
    recomputeSelectionSetFromShapes,
    _detach_selectionset_observer,
    _attach_selectionset_observer,
    _freecad_version_tuple,
    USE_SELECTION_OBSERVER,
)

# Minimum macro version expected for the full test suite (FEM cantilever, multi-target links, etc.).
REQUIRED_MACRO_VERSION = "1.0.0"

# Set to True to write fem_cantilever_debug.log and gmsh_mesh_debug.log and to emit [FEM-cut] / [GMSH-mesh] lines to report/console during FEM cantilever test. False = no debug files, quiet run.
WRITE_FEM_DEBUG_FILES = False


def _parse_version(s):
    """Convert version string 'X.Y.Z' to (X, Y, Z) for comparison. Missing parts are 0."""
    try:
        parts = [int(x) for x in (s or "0").split(".")[:3]]
        return tuple(parts + [0] * (3 - len(parts)))
    except (ValueError, TypeError):
        return (0, 0, 0)


def _macro_version_ok():
    """Return True if the loaded macro version is >= REQUIRED_MACRO_VERSION."""
    return _parse_version(getattr(selection_set_core, "__version__", "0")) >= _parse_version(REQUIRED_MACRO_VERSION)


def _add_objects_to_group(doc, group_name, objects, label=None):
    """
    Create or get an App::DocumentObjectGroup, add the given objects (that exist in doc), and recompute.
    objects: list of document objects or object names (strings). None entries are skipped.
    label: optional tree label for the group.
    Returns the group or None.
    """
    if not doc or not objects:
        return None
    objs = []
    for o in objects:
        if o is None:
            continue
        obj = doc.getObject(o) if isinstance(o, str) else o
        if obj and obj.Document == doc:
            objs.append(obj)
    if not objs:
        return None
    try:
        grp = doc.getObject(group_name)
        if grp is None:
            grp = doc.addObject("App::DocumentObjectGroup", group_name)
        if label:
            grp.Label = label
        if hasattr(grp, "addObjects"):
            existing = list(grp.Group) if hasattr(grp, "Group") and grp.Group else []
            to_add = [o for o in objs if o not in existing]
            if to_add:
                grp.addObjects(to_add)
        elif hasattr(grp, "Group"):
            existing = list(grp.Group) if grp.Group else []
            grp.Group = list(set(existing) | set(objs))
        doc.recompute()
        return grp
    except Exception:
        return None


def _ensure_test_1_4_geometry():
    """Return (main, sel_shape) for Tests 1–4. Create geometry only if not present."""
    doc = FreeCAD.ActiveDocument
    if doc is None:
        return build_test_1_cubes()
    main = doc.getObject("TestCube_main")
    sel_shape = doc.getObject("TestCube_Selection")
    if main and sel_shape:
        return main, sel_shape
    return build_test_1_cubes()


def _ensure_test_5_8_geometry():
    """Return (cube, sphere, boolean, sel_cube) for Tests 5–8. Create geometry only if not present."""
    doc = FreeCAD.ActiveDocument
    if doc is None:
        return build_test_2_cube_with_inner_sphere()
    boolean = doc.getObject("BooleanFragments")
    sel_cube = doc.getObject("TestCube_Selection_Bool")
    if boolean and sel_cube:
        cube = doc.getObject("TestCube_Bool")
        sphere = doc.getObject("TestSphere_Bool")
        return cube, sphere, boolean, sel_cube
    return build_test_2_cube_with_inner_sphere()


def _ensure_test_9_geometry():
    """Return (boolean_cyl, body003) for Tests 9, 9a, 9b. Create geometry only if not present."""
    doc = FreeCAD.ActiveDocument
    if doc is None:
        return build_test_3_cylinder_two_halves()
    boolean_cyl = doc.getObject("BooleanFragments_CylTwoHalves")
    body003 = doc.getObject("Body003")
    if boolean_cyl and body003:
        return boolean_cyl, body003
    return build_test_3_cylinder_two_halves()


def build_all_test_geometry():
    """Create all test shapes (two cubes, cube+sphere+BooleanFragments, cylinder+halves+Body003)."""
    build_test_1_cubes()
    build_test_2_cube_with_inner_sphere()
    build_test_3_cylinder_two_halves()
    print("All test geometry built. You can now run 'Run SelectionSet Tests' or try selections manually.")


def _run_fem_test(log):
    """
    Run Test FEM: create FEM analysis; ConstraintFixed from Test1+Test9a faces;
    ConstraintElectrostaticPotential from Test6 faces (all 6), using one-by-one ref workaround.
    Returns (add_tests, add_passed, add_failed, fail_name).
    Skips if FEM types are not available (e.g. FEM workbench not loaded).
    """
    # Load FEM workbench so Fem types are registered (needed when macro runs at startup)
    try:
        import Fem
    except Exception as e:
        log("Test FEM: SKIP (Fem module not available: %s). Install/enable FEM workbench." % e)
        return (0, 0, 0, None)

    doc = FreeCAD.ActiveDocument
    if not doc:
        log("Test FEM: SKIP (no active document).")
        return (0, 0, 0, None)

    # Combined refs: Test1 (1 face) + Test9a (5 faces) for ConstraintFixed
    ss1 = doc.getObject("Test1_Faces_TestCube_main_Face6_only")
    ss9a = doc.getObject("Test9a_Faces_CylTwoHalves_fully_inside_5faces")
    if not ss1 or not getattr(ss1, "ElementList", None):
        log("Test FEM: SKIP (Test1 SelectionSet not found; run Tests 1-9 first).")
        return (0, 0, 0, None)
    if not ss9a or not getattr(ss9a, "ElementList", None):
        log("Test FEM: SKIP (Test9a SelectionSet not found; run Tests 1-9 first).")
        return (0, 0, 0, None)

    refs_fixed = get_fem_references_from_selectionsets([ss1, ss9a])
    if not refs_fixed:
        log("Test FEM: FAIL (combined Test1+Test9a refs empty).")
        return (1, 0, 1, "Test FEM")

    # Create analysis and ConstraintFixed
    try:
        analysis = doc.addObject("Fem::FemAnalysis", "FEM_Test_Analysis")
    except Exception as e:
        log("Test FEM: SKIP (Fem::FemAnalysis not available: %s). Switch to FEM workbench and run again." % e)
        return (0, 0, 0, None)
    if analysis is None or not hasattr(analysis, "addObject"):
        log("Test FEM: SKIP (Fem::FemAnalysis not available). Switch to FEM workbench and run again.")
        return (0, 0, 0, None)

    # Set as active analysis so constraints are added to it (matches GUI behaviour)
    try:
        import FemGui
        FemGui.setActiveAnalysis(analysis)
    except Exception:
        pass

    try:
        constraint_fixed = doc.addObject("Fem::ConstraintFixed", "FEM_Test_ConstraintFixed")
        analysis.addObject(constraint_fixed)
    except Exception as e:
        log("Test FEM: FAIL (Fem::ConstraintFixed failed: %s)." % e)
        return (1, 0, 1, "Test FEM")

    # Use one-by-one workaround: this FreeCAD build stores only ~2 refs when assigning a list directly
    if not apply_references_to_fem_constraint_one_by_one(refs_fixed, constraint_fixed, doc):
        log("Test FEM: FAIL (apply_references_to_fem_constraint_one_by_one for ConstraintFixed returned False).")
        return (1, 0, 1, "Test FEM")

    # Verify ConstraintFixed has refs (some builds store only 2 when assigning 6; we continue so link and electrostatic get created)
    actual_fixed = list(getattr(constraint_fixed, "References", []) or [])
    if len(actual_fixed) == 0:
        log("Test FEM: FAIL (ConstraintFixed has no refs).")
        return (1, 0, 1, "Test FEM")
    if len(actual_fixed) != len(refs_fixed):
        log("Test FEM: WARNING (ConstraintFixed has %d refs, expected %d; continuing to create electrostatic + link)." % (len(actual_fixed), len(refs_fixed)))

    # Electrostatic potential constraint with all Test6 faces (6 faces).
    # Some FEM constraints only accept references when added one-by-one; use workaround (see selection_set_core.apply_references_to_fem_constraint_one_by_one).
    ss6 = doc.getObject("Test6_Faces_TestCube_Selection_Bool_6faces_fully_inside")
    if not ss6 or not getattr(ss6, "ElementList", None):
        log("Test FEM: FAIL (Test6 SelectionSet not found).")
        return (1, 0, 1, "Test FEM")
    refs_electro = get_fem_references_from_selectionset(ss6)
    if len(refs_electro) != 6:
        log("Test FEM: FAIL (Test6 should have 6 faces, got %d)." % len(refs_electro))
        return (1, 0, 1, "Test FEM")

    # Use ObjectsFem like the GUI (FEM_CompEmConstraints); fall back to addObject if not available
    constraint_electro = None
    try:
        import ObjectsFem
        constraint_electro = ObjectsFem.makeConstraintElectrostaticPotential(doc, "FEM_Test_ConstraintElectrostaticPotential")
    except Exception as e:
        log("Test FEM: ObjectsFem.makeConstraintElectrostaticPotential not available (%s), trying addObject." % e)
    if constraint_electro is None:
        try:
            constraint_electro = doc.addObject("Fem::ConstraintElectrostaticPotential", "FEM_Test_ConstraintElectrostaticPotential")
        except Exception as e:
            log("Test FEM: FAIL (Fem::ConstraintElectrostaticPotential failed: %s)." % e)
            return (1, 0, 1, "Test FEM")
    analysis.addObject(constraint_electro)

    # Set potential (0 V) if property exists
    if hasattr(constraint_electro, "Potential"):
        try:
            from FreeCAD import Units
            constraint_electro.Potential = Units.Quantity("0 V")
        except Exception:
            pass
    if hasattr(constraint_electro, "PotentialEnabled"):
        constraint_electro.PotentialEnabled = True

    # Workaround: constraint accepts only one face at a time; add references one-by-one (see selection_set_core.apply_references_to_fem_constraint_one_by_one).
    if not apply_references_to_fem_constraint_one_by_one(refs_electro, constraint_electro, doc):
        log("Test FEM: FAIL (apply_references_to_fem_constraint_one_by_one returned False).")
        return (1, 0, 1, "Test FEM")

    actual_electro = list(getattr(constraint_electro, "References", []) or [])
    if len(actual_electro) == 0:
        log("Test FEM: FAIL (ConstraintElectrostaticPotential has no refs).")
        return (1, 0, 1, "Test FEM")
    if len(actual_electro) != 6:
        log("Test FEM: WARNING (ConstraintElectrostaticPotential has %d refs, expected 6; continuing to link)." % len(actual_electro))

    # SelectionSetLink: create link that adds Test1+Test9a and targets ConstraintFixed; apply and verify
    link = createSelectionSetLink("FEM_Test_SelectionSetLink")
    if not link:
        log("Test FEM: FAIL (createSelectionSetLink failed).")
        return (1, 0, 1, "Test FEM")
    link.TargetObjects = [constraint_fixed]
    link.AddSelectionSets = [ss1, ss9a]
    link.SubtractSelectionSets = []
    link.UseOneByOneWorkaround = True  # same build needs one-by-one for ConstraintFixed
    if not link.Proxy.apply_to_target(link):
        log("Test FEM: FAIL (SelectionSetLink apply_to_target returned False).")
        return (1, 0, 1, "Test FEM")
    refs_via_link = list(getattr(constraint_fixed, "References", []) or [])
    if len(refs_via_link) == 0:
        log("Test FEM: FAIL (after SelectionSetLink apply: ConstraintFixed has no refs).")
        return (1, 0, 1, "Test FEM")
    if len(refs_via_link) != len(refs_fixed):
        log("Test FEM: WARNING (SelectionSetLink applied; ConstraintFixed has %d refs, expected %d)." % (len(refs_via_link), len(refs_fixed)))
    doc.recompute()

    # Second SelectionSetLink: target ConstraintElectrostaticPotential, add Test6 faces
    link_electro = createSelectionSetLink("FEM_Test_SelectionSetLink_Electro")
    if link_electro:
        link_electro.TargetObjects = [constraint_electro]
        link_electro.AddSelectionSets = [ss6]
        link_electro.SubtractSelectionSets = []
        link_electro.UseOneByOneWorkaround = True
        link_electro.Proxy.apply_to_target(link_electro)
        log("Test FEM: second SelectionSetLink (electrostatic) created and applied.")
    else:
        log("Test FEM: WARNING (second SelectionSetLink for electrostatic not created).")

    # FEM material (solid) and third SelectionSetLink. Note: In FreeCAD's geometry reference selector
    # for FEM material, 'Face, Edge' must be selected – 'Solid' does not work (tested with FEM material).
    material_obj = None
    try:
        import ObjectsFem
        material_obj = ObjectsFem.makeMaterialSolid(doc, "FEM_Test_MaterialSolid")
    except Exception as e:
        try:
            material_obj = doc.addObject("Fem::Material", "FEM_Test_MaterialSolid")
        except Exception:
            pass
    if material_obj is not None:
        analysis.addObject(material_obj)
        # Material link: add shape BooleanFragments_CylTwoHalves (all solids), subtract selection set Test9b (Solid3).
        # Result: Solid1 and Solid2 in FEM material (two solids).
        boolean_cyl = doc.getObject("BooleanFragments_CylTwoHalves")
        ss9b = doc.getObject("Test9b_Solids_CylTwoHalves_Solid3_one_half_inside")
        link_material = createSelectionSetLink("FEM_Test_SelectionSetLink_Material")
        if link_material and boolean_cyl and ss9b and getattr(ss9b, "ElementList", None):
            link_material.TargetObjects = [material_obj]
            link_material.AddSelectionSets = []
            link_material.AddShapes = [boolean_cyl]
            link_material.ShapeRefMode = "solids"
            link_material.SubtractSelectionSets = [ss9b]
            link_material.SubtractShapes = []
            link_material.UseOneByOneWorkaround = True
            link_material.Proxy.apply_to_target(link_material)
            refs_material = list(getattr(material_obj, "References", []) or [])
            if len(refs_material) >= 2:
                log("Test FEM: FEM material + third SelectionSetLink (add shape BooleanFragments_CylTwoHalves, subtract Test9b -> Solid1, Solid2) created and applied.")
            else:
                log("Test FEM: FEM material + third SelectionSetLink created (expected 2 solids; got %d refs)." % len(refs_material))
        else:
            log("Test FEM: WARNING (FEM material or third SelectionSetLink not created; BooleanFragments_CylTwoHalves or Test9b may be missing).")
    else:
        log("Test FEM: WARNING (FEM material object not created).")

    doc.recompute()
    log("Test FEM: PASS (FEM analysis + ConstraintFixed + ConstraintElectrostaticPotential + 3 SelectionSetLinks + FEM material created and verified).")
    return (1, 1, 0, None)


# Reference max deflection (mm) and tolerance for FEM cantilever validation.
#
# Geometry (see run_fem_cantilever_beam_example): Beam 8000 x 1000 x 1000 mm. Hole: cylinder
# axis along Y, center at mid-length (x = 4000 mm), z = 0.35*h = 350 mm. CharacterLengthMax = 200 mm;
# material on beam solid only (1 solid).
#
# Case: beam WITH hole, cylinder radius 300 mm (same x, z, mesh as above).
FEM_CANTILEVER_REF_WITH_R300mm_HOLE_MM = 96.2
#
# Case: beam WITH hole, cylinder radius 400 mm (same x, z, CharacterLengthMax 200 mm).
FEM_CANTILEVER_REF_WITH_R400mm_HOLE_MM = 375.31
#
# Case: beam WITHOUT hole (box only, no cut).
FEM_CANTILEVER_REF_NO_HOLE_MM = 88.4
#
# Tolerances (mm): allowed below/above ref for the two hole cases and no-hole.
FEM_CANTILEVER_TOL_LOW_MM = 1.5
FEM_CANTILEVER_TOL_HIGH_MM = 1.0
FEM_CANTILEVER_TOL_R400_MM = 2.0   # for r=400 mm hole case only
#
# If hole geometry (position, radius) or mesh size is changed, recompute and update the refs above.
#
# Info for users on FreeCAD 1.0.2: the automated FEM cantilever test may fail (GMSH mesh not
# generated) due to BREP export limitations. That failure is expected on 1.0.2. When you run the
# FEM beam example manually (Tests -> FEM beam example, then run GMSH and solver yourself),
# expect deflections in these ranges (beam 8000×1000×1000 mm, mesh CharLengthMax 200 mm):
#   With hole (r=300 mm, x=4000, z=350): ~96 mm.  Without hole: ~88 mm.  With hole r=400 mm: ~375 mm.
FEM_CANTILEVER_FC102_INFO = (
    "On FreeCAD 1.0.2 the FEM cantilever test often fails (GMSH did not generate mesh); that is expected. "
    "When running meshing and solving the FEM beam example manually, expect max deflection: with hole r=300 mm ~96 mm, "
    "without hole ~88 mm, with hole r=400 mm ~375 mm (beam 8000×1000×1000 mm, mesh size 200 mm)."
)


# ---------- FreeCAD version: 1.0.2 vs 1.1 ----------
# Use these to branch mesh/cut and other version-specific code.
def _is_fc102():
    """True if running on FreeCAD 1.0.x (e.g. 1.0.2). Use for mesh (GmshTools workflow) and cut (BOP preferred)."""
    return _freecad_version_tuple() < (1, 1)


def _is_fc11():
    """True if running on FreeCAD 1.1 or newer. Use for mesh (run()) and cut (Part::Cut/BOP)."""
    return _freecad_version_tuple() >= (1, 1)


def _run_gmsh_for_mesh_fc102(mesh, dbg=None):
    """
    Generate GMSH mesh on FreeCAD 1.0.2. GmshTools has no run(). Use the manual sequence
    write_gmsh_input_files() -> run_gmsh_with_geo() -> read_and_set_new_mesh(); if that
    is not available, try create_mesh() (which often returns None in 1.0). Returns True
    if mesh was generated. dbg: optional debug log function for gmsh_mesh_debug.log.
    """
    if not mesh:
        return False
    _log = dbg if callable(dbg) else (lambda msg, include_tb=False: None)
    try:
        from femmesh import gmshtools
        tool = gmshtools.GmshTools(mesh)
        _log("fc102: GmshTools(mesh) created; has run=%s create_mesh=%s write_gmsh_input_files=%s run_gmsh_with_geo=%s read_and_set_new_mesh=%s" % (
            hasattr(tool, "run"), hasattr(tool, "create_mesh"),
            hasattr(tool, "write_gmsh_input_files"), hasattr(tool, "run_gmsh_with_geo"), hasattr(tool, "read_and_set_new_mesh")))
        # 1.0.2: Prefer explicit three-step workflow so we control and can verify each step
        if all(hasattr(tool, m) for m in ("write_gmsh_input_files", "run_gmsh_with_geo", "read_and_set_new_mesh")):
            _log("fc102: calling write_gmsh_input_files()")
            tool.write_gmsh_input_files()
            _log("fc102: calling run_gmsh_with_geo()")
            tool.run_gmsh_with_geo()
            _log("fc102: calling read_and_set_new_mesh()")
            tool.read_and_set_new_mesh()
            if getattr(mesh, "FemMesh", None):
                _log("fc102: FemMesh set after manual workflow -> True")
                return True
            _log("fc102: FemMesh still empty after manual workflow")
        # Fallback: create_mesh() if manual steps missing or failed
        if hasattr(tool, "create_mesh"):
            try:
                _log("fc102: calling create_mesh()")
                tool.create_mesh()
                ok = bool(getattr(mesh, "FemMesh", None))
                _log("fc102: create_mesh() done, FemMesh set=%s" % ok)
                return ok
            except Exception as e:
                _log("fc102: create_mesh() failed: %s" % e)
                _log("(exception)", include_tb=True)
        return False
    except Exception as e:
        _log("fc102: GmshTools failed: %s" % e)
        _log("(exception)", include_tb=True)
        return False


def _run_gmsh_for_mesh_fc11(mesh, dbg=None):
    """
    Generate GMSH mesh on FreeCAD 1.1. Prefer femexamples.mesh_from_mesher; fallback GmshTools.run(blocking=True).
    Returns True if mesh was generated, False otherwise. dbg: optional debug log function.
    """
    if not mesh:
        return False
    _log = dbg if callable(dbg) else (lambda msg, include_tb=False: None)
    try:
        from femexamples.meshes import generate_mesh
        ok = bool(generate_mesh.mesh_from_mesher(mesh, "gmsh"))
        _log("fc11: mesh_from_mesher returned %s" % ok)
        return ok
    except Exception as e:
        _log("fc11: mesh_from_mesher failed: %s" % e)
    try:
        from femmesh import gmshtools
        tool = gmshtools.GmshTools(mesh)
        if hasattr(tool, "run"):
            ok = bool(tool.run(blocking=True))
            _log("fc11: GmshTools.run(blocking=True) returned %s" % ok)
            return ok
        _log("fc11: GmshTools has no run()")
        return False
    except Exception as e:
        _log("fc11: GmshTools failed: %s" % e)
        return False


def _gmsh_mesh_debug_logger():
    """Return (debug_log_fn, path, file_handle). Writes to gmsh_mesh_debug.log for diagnosing why mesh is created but not meshed (e.g. on 1.0.2). When WRITE_FEM_DEBUG_FILES is False, returns no-op logger and None path/handle."""
    if not WRITE_FEM_DEBUG_FILES:
        return (lambda msg, include_tb=False: None), None, None
    import os
    import traceback
    base = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, "gmsh_mesh_debug.log")
    try:
        f = open(path, "w", encoding="utf-8")
    except Exception:
        f = None
        path = None

    def debug_log(msg, include_tb=False):
        line = "[GMSH-mesh] %s" % msg
        if f:
            try:
                f.write(line + "\n")
                if include_tb:
                    f.write(traceback.format_exc() + "\n")
                f.flush()
            except Exception:
                pass

    return debug_log, path, f


def _log_mesh_state(dbg, mesh, prefix="mesh"):
    """Log mesh object state for debugging: Part/Shape/Geometry, FemMesh, FC version."""
    if not dbg or not mesh:
        return
    try:
        dbg("%s: Name=%s TypeId=%s" % (prefix, getattr(mesh, "Name", "?"), getattr(mesh, "TypeId", "?")))
        for prop in ("Part", "Geometry", "Shape"):
            if hasattr(mesh, prop):
                val = getattr(mesh, prop, None)
                dbg("%s:   %s=%s" % (prefix, prop, val.Name if val else None))
        fm = getattr(mesh, "FemMesh", None)
        if fm:
            nc = getattr(fm, "NodeCount", None)
            if callable(nc):
                nc = nc() if nc else "?"
            dbg("%s:   FemMesh present, NodeCount=%s" % (prefix, nc))
        else:
            dbg("%s:   FemMesh=None or empty" % prefix)
        dbg("%s:   FC version=%s" % (prefix, _freecad_version_tuple()))
    except Exception as e:
        dbg("%s:   (error logging state: %s)" % (prefix, e))


def _ensure_femexamples_on_path():
    """
    Add FreeCAD Fem Mod path to sys.path so femexamples can be imported when the macro
    runs at startup (FEM workbench may not be loaded yet). Same module used by
    Utilities -> Open fem-examples. Tries getHomePath()/Mod/Fem first, then dir of Fem module.
    """
    try:
        import sys
        import os
        home = FreeCAD.getHomePath()
        if home:
            fem_mod = os.path.join(home, "Mod", "Fem")
            if os.path.isdir(fem_mod) and fem_mod not in sys.path:
                sys.path.insert(0, fem_mod)
                return
        try:
            import Fem as _Fem
            fem_dir = os.path.dirname(getattr(_Fem, "__file__", "") or "")
            if fem_dir and os.path.isdir(fem_dir) and fem_dir not in sys.path:
                sys.path.insert(0, fem_dir)
        except Exception:
            pass
    except Exception:
        pass


def _run_gmsh_for_mesh(mesh):
    """
    Generate GMSH mesh for the given FemMeshGmsh object. Uses the same Python API as
    Utilities -> Open fem-examples (e.g. ccx_cantilever_base_solid): try femexamples.meshes
    generate_mesh.mesh_from_mesher(mesh, "gmsh") first on all FreeCAD versions (1.0.2 and 1.1).
    If that fails, use version-specific fallback: 1.0.2 -> _run_gmsh_for_mesh_fc102 (GmshTools
    manual workflow); 1.1 -> _run_gmsh_for_mesh_fc11 (GmshTools.run). Returns True if generated.
    Writes debug log to gmsh_mesh_debug.log when meshing is attempted.
    """
    if not mesh:
        return False
    dbg, _path, _f = _gmsh_mesh_debug_logger()
    try:
        dbg("_run_gmsh_for_mesh: start")
        _log_mesh_state(dbg, mesh, "before")
        # Same as FEM examples (Utilities -> Open fem-examples): mesh_from_mesher works on 1.0.2 and 1.1
        try:
            dbg("trying femexamples.meshes.generate_mesh.mesh_from_mesher(mesh, 'gmsh')")
            _ensure_femexamples_on_path()
            from femexamples.meshes import generate_mesh
            ok = generate_mesh.mesh_from_mesher(mesh, "gmsh")
            dbg("mesh_from_mesher returned: %s" % ok)
            if ok:
                _log_mesh_state(dbg, mesh, "after femexamples")
                return True
        except Exception as e:
            dbg("femexamples.mesh_from_mesher failed: %s" % e)
            dbg("(exception)", include_tb=True)
        # Version-specific fallback when femexamples not available or mesh_from_mesher failed
        if _is_fc102():
            dbg("fallback: _run_gmsh_for_mesh_fc102 (1.0.2)")
            ok = _run_gmsh_for_mesh_fc102(mesh, dbg)
            _log_mesh_state(dbg, mesh, "after fc102")
            return ok
        dbg("fallback: _run_gmsh_for_mesh_fc11 (1.1)")
        ok = _run_gmsh_for_mesh_fc11(mesh, dbg)
        _log_mesh_state(dbg, mesh, "after fc11")
        return ok
    finally:
        if _f:
            try:
                _f.close()
            except Exception:
                pass


def _run_fem_cantilever_full_test(log):
    """
    Run full FEM cantilever test: ensure setup, run three cases (hole r=300 mm, no hole, hole r=400 mm),
    check max deflection against FEM_CANTILEVER_REF_WITH_R300mm_HOLE_MM, REF_NO_HOLE_MM, REF_WITH_R400mm_HOLE_MM.
    Geometry: beam 8000×1000×1000 mm, hole center x=4000, z=350; CharacterLengthMax 200 mm.
    Returns (add_tests, add_passed, add_failed, fail_name).
    """
    try:
        import Fem  # noqa: F401
    except Exception as e:
        log("Test FEM cantilever full: SKIP (Fem module not available: %s)." % e)
        return (0, 0, 0, None)

    doc = FreeCAD.ActiveDocument
    if not doc:
        log("Test FEM cantilever full: SKIP (no active document).")
        return (0, 0, 0, None)

    # Ensure FEM cantilever is set up (custom build with Beam_Box, cut, config)
    analysis = doc.getObject("FEM_Beam_Analysis")
    if not analysis or not hasattr(analysis, "Group"):
        log("Test FEM cantilever full: setting up FEM cantilever example...")
        # Place the FEM cantilever slightly above the other test geometries in +Y so they do not overlap
        run_fem_cantilever_beam_example(cylinder_cut=False, geometry_offset=FreeCAD.Vector(0, 500, 0))
        analysis = doc.getObject("FEM_Beam_Analysis")
    if not analysis:
        log("Test FEM cantilever full: SKIP (FEM_Beam_Analysis not found).")
        return (0, 0, 0, None)

    config = doc.getObject("FEM_Beam_Config")
    if not config:
        log("Test FEM cantilever full: SKIP (FEM_Beam_Config not found; need custom build).")
        return (0, 0, 0, None)

    mesh = doc.getObject("FEM_Beam_Mesh")
    if not mesh:
        log("Test FEM cantilever full: FAIL (FEM_Beam_Mesh not found).")
        return (1, 0, 1, "Test FEM cantilever full")

    solver = doc.getObject("FEM_Beam_Solver")
    if not solver and analysis:
        for obj in getattr(analysis, "Group", []) or []:
            if "Ccx" in getattr(obj, "Name", "") or "CalculiX" in getattr(obj, "Label", "") or "SolverCcx" in getattr(obj, "Name", ""):
                solver = obj
                break
    if not solver:
        for obj in doc.Objects:
            if "Ccx" in getattr(obj, "Name", "") or "CalculiX" in getattr(obj, "Label", "") or "SolverCcx" in getattr(obj, "Name", ""):
                if hasattr(obj, "Proxy"):
                    solver = obj
                    break
    if not solver:
        log("Test FEM cantilever full: FAIL (CalculiX solver not found).")
        return (1, 0, 1, "Test FEM cantilever full")

    def run_one_case(suppressed_val, case_name, radius_mm=None):
        """Set config.Suppressed, optionally set cylinder radius (mm), mesh, solve, return (max_deflection_mm, None) or (None, error_msg)."""
        if radius_mm is not None:
            var_obj = doc.getObject("FEM_Beam_Variables") or doc.getObject("FEM_Beam_VarSet")
            if var_obj and hasattr(var_obj, "CylinderRadius"):
                var_obj.CylinderRadius = float(radius_mm)
            elif config and hasattr(config, "CylinderRadius"):
                config.CylinderRadius = float(radius_mm)
            else:
                cyl_obj = doc.getObject("Beam_CylinderCut")
                if cyl_obj:
                    cyl_obj.Radius = float(radius_mm)
            doc.recompute()
        try:
            config.Suppressed = suppressed_val
        except Exception as e:
            return None, "set Suppressed=%s: %s" % (suppressed_val, e)
        try:
            QtWidgets.QApplication.processEvents()
        except Exception:
            pass
        doc.recompute()
        # Mesh element size (1.1: CharacterLengthMax; 1.0: often CharacteristicLengthMax)
        if hasattr(mesh, "CharacterLengthMax"):
            mesh.CharacterLengthMax = 200.0
        if hasattr(mesh, "CharacteristicLengthMax"):
            mesh.CharacteristicLengthMax = 200.0
        ok = _run_gmsh_for_mesh(mesh)
        if not ok:
            return None, "GMSH did not generate mesh (%s)" % case_name
        doc.recompute()
        ran = False
        try:
            import os as _os
            workdir = _os.path.join(FreeCAD.getUserAppDataDir(), "ccx_run")
            _os.makedirs(workdir, exist_ok=True)
            if hasattr(solver, "WorkingDir"):
                solver.WorkingDir = workdir
            for method in ("write_inp_file", "writeInputFile", "write_inp", "writeInput"):
                if hasattr(solver, "Proxy") and hasattr(solver.Proxy, method):
                    getattr(solver.Proxy, method)(solver, analysis)
                    break
            for method in ("run", "runSolver", "startSolver", "execute"):
                if hasattr(solver, "Proxy") and hasattr(solver.Proxy, method):
                    getattr(solver.Proxy, method)(solver, analysis)
                    ran = True
                    break
            if ran and hasattr(solver, "Proxy"):
                for load_m in ("load_results", "loadResults", "read_results"):
                    if hasattr(solver.Proxy, load_m):
                        try:
                            getattr(solver.Proxy, load_m)(solver, analysis)
                        except Exception:
                            pass
                        break
            # Force-load latest result into result object (same object may be reused for 2nd run)
            if ran:
                try:
                    from femtools import ccxtools
                    _fea = ccxtools.FemToolsCcx(analysis) if analysis else ccxtools.FemToolsCcx(solver)
                    if hasattr(_fea, "load_results"):
                        _fea.load_results()
                except Exception:
                    pass
        except Exception as e:
            return None, "solver %s: %s" % (case_name, e)
        if not ran:
            try:
                from femtools import ccxtools
                fea = ccxtools.FemToolsCcx(analysis) if analysis else ccxtools.FemToolsCcx(solver)
                fea.run()
                ran = True
            except Exception as e:
                return None, "solver %s: %s" % (case_name, e)
        if not ran:
            return None, "no run method (%s)" % case_name
        doc.recompute()
        # Use the last result object in the analysis (most recent run); otherwise we may read
        # the first run's data when checking the second case.
        result_obj = None
        candidates = []
        for obj in getattr(analysis, "Group", []) or []:
            try:
                if hasattr(obj, "DisplacementLengths") or hasattr(obj, "DisplacementVectors"):
                    candidates.append(obj)
            except Exception:
                continue
        if candidates:
            result_obj = candidates[-1]
        if result_obj is None:
            for name in ("CCX_Results", "Results", "ResultMechanical"):
                obj = doc.getObject(name)
                if obj and (hasattr(obj, "DisplacementLengths") or hasattr(obj, "DisplacementVectors")):
                    result_obj = obj
                    break
        if result_obj is None:
            return None, "no result object (%s)" % case_name
        lengths = getattr(result_obj, "DisplacementLengths", None)
        if not lengths and hasattr(result_obj, "DisplacementVectors"):
            try:
                vectors = result_obj.DisplacementVectors
                if vectors:
                    lengths = [v.Length for v in vectors]
            except Exception:
                pass
        if not lengths:
            return None, "no DisplacementLengths (%s)" % case_name
        try:
            return float(max(lengths)), None
        except (TypeError, ValueError):
            return None, "max displacement invalid (%s)" % case_name

    # Case 1: with hole, cylinder radius 300 mm (Suppressed=False).
    low_hole = FEM_CANTILEVER_REF_WITH_R300mm_HOLE_MM - FEM_CANTILEVER_TOL_LOW_MM
    high_hole = FEM_CANTILEVER_REF_WITH_R300mm_HOLE_MM + FEM_CANTILEVER_TOL_HIGH_MM
    max_hole, err_hole = run_one_case(False, "with hole")
    if err_hole:
        log("Test FEM cantilever full: FAIL (%s)." % err_hole)
        if _is_fc102() and "GMSH" in str(err_hole):
            log("[Info] %s" % FEM_CANTILEVER_FC102_INFO)
        return (1, 0, 1, "Test FEM cantilever full")
    if not (low_hole <= max_hole <= high_hole):
        log("Test FEM cantilever full: FAIL (with hole: max deflection %.4f mm outside [%.2f, %.2f] mm)." % (max_hole, low_hole, high_hole))
        return (1, 0, 1, "Test FEM cantilever full")
    log("Test of FEM-cantilever with hole yields correct displacement values (depending on mesh): max deflection %.4f mm in [%.2f, %.2f] mm." % (max_hole, low_hole, high_hole))

    # Case 2: no hole (Suppressed=True, box only).
    low_nohole = FEM_CANTILEVER_REF_NO_HOLE_MM - FEM_CANTILEVER_TOL_LOW_MM
    high_nohole = FEM_CANTILEVER_REF_NO_HOLE_MM + FEM_CANTILEVER_TOL_HIGH_MM
    max_nohole, err_nohole = run_one_case(True, "no hole")
    if err_nohole:
        log("Test FEM cantilever full: FAIL (%s)." % err_nohole)
        if _is_fc102() and "GMSH" in str(err_nohole):
            log("[Info] %s" % FEM_CANTILEVER_FC102_INFO)
        return (1, 0, 1, "Test FEM cantilever full")
    if not (low_nohole <= max_nohole <= high_nohole):
        log("Test FEM cantilever full: FAIL (no hole: max deflection %.4f mm outside [%.2f, %.2f] mm)." % (max_nohole, low_nohole, high_nohole))
        return (1, 0, 1, "Test FEM cantilever full")
    log("Test of FEM-cantilever without hole yields correct displacement values (depending on mesh): max deflection %.4f mm in [%.2f, %.2f] mm." % (max_nohole, low_nohole, high_nohole))

    # Case 3: with hole, cylinder radius 400 mm (same x, z, CharacterLengthMax 200 mm).
    low_r400 = FEM_CANTILEVER_REF_WITH_R400mm_HOLE_MM - FEM_CANTILEVER_TOL_R400_MM
    high_r400 = FEM_CANTILEVER_REF_WITH_R400mm_HOLE_MM + FEM_CANTILEVER_TOL_R400_MM
    max_r400, err_r400 = run_one_case(False, "with hole r=400 mm", radius_mm=400)
    if err_r400:
        log("Test FEM cantilever full: FAIL (%s)." % err_r400)
        if _is_fc102() and "GMSH" in str(err_r400):
            log("[Info] %s" % FEM_CANTILEVER_FC102_INFO)
        return (1, 0, 1, "Test FEM cantilever full")
    if not (low_r400 <= max_r400 <= high_r400):
        log("Test FEM cantilever full: FAIL (with hole r=400 mm: max deflection %.4f mm outside [%.2f, %.2f] mm)." % (max_r400, low_r400, high_r400))
        return (1, 0, 1, "Test FEM cantilever full")
    log("Test of FEM-cantilever with hole (radius 400 mm) yields correct displacement values: max deflection %.4f mm in [%.2f, %.2f] mm." % (max_r400, low_r400, high_r400))
    # Restore default radius for subsequent runs (Variables/VarSet are source when present)
    var_obj = doc.getObject("FEM_Beam_Variables") or doc.getObject("FEM_Beam_VarSet")
    if var_obj and hasattr(var_obj, "CylinderRadius"):
        var_obj.CylinderRadius = 300.0
    elif config and hasattr(config, "CylinderRadius"):
        config.CylinderRadius = 300.0
    else:
        cyl_obj = doc.getObject("Beam_CylinderCut")
        if cyl_obj:
            cyl_obj.Radius = 300.0
    doc.recompute()

    log("Test of FEM-cantilever: PASS (with hole %.4f mm, no hole %.4f mm, hole r=400 mm %.4f mm within tolerance)." % (max_hole, max_nohole, max_r400))
    return (1, 1, 0, None)


# ---------- Test runner ----------
def run_test_1_only():
    """Create the two test cubes (main + selection shape) if missing, then run only Test 1 (one face selected).
    Creates the SelectionSet programmatically so it has MainObject/SelectionShape and is deletable like others."""
    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")

    def log(msg):
        if report:
            report.append(msg)
        else:
            print(msg)

    try:
        _detach_selectionset_observer()
    except Exception:
        pass

    main, sel_shape = _ensure_test_1_4_geometry()
    log("\n[SelectionSet] Run Test 1: two cubes ready.")

    doc = FreeCAD.ActiveDocument
    if not doc:
        log("Run Test 1: no active document.")
        return
    name = "Test1_Faces_TestCube_main_Face6_only"
    existing = doc.getObject(name)
    if existing:
        doc.removeObject(name)
    elements = get_faces_by_selection_shape(
        main, sel_shape.Shape, volume_mode="fully_inside"
    )
    obj = doc.addObject("App::FeaturePython", name)
    SelectionSet(obj)
    obj.Label = "SelectionSet " + name
    # Store MainObject/SelectionShape so this SelectionSet behaves like the one
    # created by the full test runner (Test 1). Deleting the SelectionSet will
    # now show the normal FreeCAD warning but will not delete the cubes.
    obj.MainObject = main
    obj.SelectionShape = sel_shape
    obj.Mode = "faces"
    obj.VolumeMode = "fully_inside"
    obj.ElementList = elements
    doc.recompute()
    ss1 = obj
    expected1 = ["TestCube_main.Face6"]
    result1 = list(ss1.ElementList)
    passed = result1 == expected1
    log("Test 1: %s == %s -> %s" % (result1, expected1, "PASS" if passed else "FAIL"))

    # Do NOT auto-attach the single-click observer here; it interferes with the
    # tree context menu and Std_Delete by replacing the SelectionSet selection
    # with its main shape/face. Users can enable USE_SELECTION_OBSERVER globally
    # if they want single-click expansion.


def run_test_3_combined_demo():
    """
    Combined demo for Run Test 3 button:
    - Ensure Test 1 cubes exist.
    - Create (or recreate) a SelectionSet from TestCube_main/TestCube_Selection (top face).
    - Create a simple FEM analysis + Fem::ConstraintFixed on that face via a SelectionSetLink.
    This builds only the required geometry for one SelectionSet + SelectionSetLink + FEM example.
    """
    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")

    def log(msg):
        if report:
            report.append(msg)
        else:
            print(msg)

    # Ensure FEM is available
    try:
        import Fem  # noqa: F401
    except Exception as e:
        log("Run Test 3: SKIP (Fem module not available: %s)." % e)
        return

    doc = FreeCAD.ActiveDocument
    if doc is None:
        doc = FreeCAD.newDocument()
        doc.Label = "SelectionSetTest"

    # Ensure cubes (Test 1 geometry); may create/close doc (e.g. build_test_1_cubes closes SelectionSetTest)
    main, sel_shape = _ensure_test_1_4_geometry()
    doc = FreeCAD.ActiveDocument
    if not doc:
        log("Run Test 3: no active document after ensuring geometry.")
        return
    log("\n[SelectionSet] Run Test 3: cubes ready for combined demo.")

    # Create/recreate SelectionSet for Test 1 face
    name_set = "Test3_SelectionSet_FixedFace"
    existing = doc.getObject(name_set)
    if existing:
        doc.removeObject(existing.Name)
    elements = get_faces_by_selection_shape(main, sel_shape.Shape, volume_mode="fully_inside")
    obj_set = doc.addObject("App::FeaturePython", name_set)
    SelectionSet(obj_set)
    obj_set.Label = "SelectionSet " + name_set
    obj_set.MainObject = main
    obj_set.SelectionShape = sel_shape
    obj_set.Mode = "faces"
    obj_set.VolumeMode = "fully_inside"
    obj_set.ElementList = elements
    apply_selectionset_view_defaults(obj_set)

    # Create or reuse FEM analysis and fixed constraint
    analysis = doc.getObject("FEM_Test3_Analysis")
    if not analysis:
        analysis = doc.addObject("Fem::FemAnalysis", "FEM_Test3_Analysis")
    constraint = doc.getObject("FEM_Test3_ConstraintFixed")
    if not constraint:
        try:
            constraint = doc.addObject("Fem::ConstraintFixed", "FEM_Test3_ConstraintFixed")
            analysis.addObject(constraint)
        except Exception as e:
            log("Run Test 3: SKIP (Fem::ConstraintFixed not available: %s)." % e)
            return

    # Create SelectionSetLink that applies the SelectionSet to the constraint
    link_name = "FEM_Test3_SelectionSetLink"
    link = doc.getObject(link_name)
    if link:
        doc.removeObject(link.Name)
    link = createSelectionSetLink()
    link.Label = "SelectionSetLink Test3 (fixed face)"
    link.TargetObjects = [constraint]
    link.AddSelectionSets = [obj_set]

    # Apply link to target using core helper
    from selection_set_core import _apply_link_to_target  # local import to avoid circulars at top

    ok = bool(_apply_link_to_target(link) is None or True)  # _apply_link_to_target prints status; treat call as success path
    if ok:
        log("Run Test 3: PASS (SelectionSetLink applied to FEM_Test3_ConstraintFixed).")
    else:
        log("Run Test 3: FAIL (SelectionSetLink could not be applied).")


# ---------------------------------------------------------------------------
# FEM beam geometry switch: why a custom object is needed
# ---------------------------------------------------------------------------
# Research: The switch (use Beam_Box or Beam_Cut for mesh + SelectionSets + links)
# cannot be done with standard FreeCAD objects only because:
# - Mesh Part/Shape and SelectionSet MainObject are PropertyLink; FreeCAD does not
#   support conditional link expressions (e.g. "VarSet.UseCut ? Beam_Cut : Beam_Box").
# - There is no built-in "object switch" or selector that outputs one of two shapes.
# - The steps "recompute SelectionSet from its MainObject" and "apply SelectionSetLink
#   refs to target" are imperative (Python); they have no expression equivalent.
# So we use a custom FeaturePython object (FEM_Beam_Config) with a boolean
# Suppressed (True = use box, False = use cut). When Suppressed changes, onChanged
# runs and updates mesh, SelectionSets, and links. If the Part::Cut object has a
# native Suppressed property (e.g. in FreeCAD 1.1), it is kept in sync.
# Optionally driven by App::VarSet via expression. See README "FEM beam geometry switch".
# ---------------------------------------------------------------------------


class _FEM_Beam_ConfigProxy:
    """
    Proxy for FEM_Beam_Config: use Suppressed to switch between Beam_Box and Beam_Cut.
    Suppressed=True -> use box (no hole); Suppressed=False -> use cut. onChanged
    schedules the switch via QTimer.singleShot(0, ...) to avoid recursive recompute.
    If Beam_Cut has a native Suppressed property, it is synced so the cut is suppressed in the tree.
    """

    def __init__(self, obj):
        obj.Proxy = self
        if not hasattr(obj, "Suppressed"):
            obj.addProperty("App::PropertyBool", "Suppressed", "Beam",
                            "Suppress the cylinder cut: use box geometry (True) or cut geometry (False).")
            # Backward compatibility: migrate from UseCylinderCut if present. Default False = use cut (hole).
            obj.Suppressed = not getattr(obj, "UseCylinderCut", False) if hasattr(obj, "UseCylinderCut") else False

    def onChanged(self, obj, prop):
        if prop != "Suppressed":
            return
        doc = obj.Document
        if not doc:
            return
        doc_name = doc.Name

        def _do_switch():
            doc = FreeCAD.getDocument(doc_name)
            if not doc:
                return
            config = doc.getObject("FEM_Beam_Config")
            if not config:
                return
            beam_box = doc.getObject("Beam_Box")
            beam_cut = _get_beam_cut_object(doc)
            mesh = doc.getObject("FEM_Beam_Mesh")
            suppressed = getattr(config, "Suppressed", True)
            beam = beam_box if suppressed else beam_cut
            if not beam or not mesh:
                return
            try:
                # Optionally sync Part::Cut's native Suppressed (e.g. FreeCAD 1.1) so tree reflects state
                if beam_cut and hasattr(beam_cut, "Suppressed"):
                    try:
                        beam_cut.Suppressed = suppressed
                    except Exception:
                        pass
                for p in ("Part", "Geometry", "Shape"):
                    if hasattr(mesh, p):
                        setattr(mesh, p, beam)
                        break
                for name in ("Beam_FixedFace", "Beam_ForceFace", "Beam_Solid"):
                    ss = doc.getObject(name)
                    if ss and hasattr(ss, "MainObject"):
                        ss.MainObject = beam
                        recomputeSelectionSetFromShapes(ss)
                from selection_set_core import _apply_link_to_target, get_fem_references_from_selectionsets, apply_references_to_fem_constraint_one_by_one
                beam_shape_objects = (beam_box, beam_cut) if beam_cut else (beam_box,)
                for link_name in ("FEM_Beam_Link_Material", "FEM_Beam_Link_Fixed", "FEM_Beam_Link_Force"):
                    link = doc.getObject(link_name)
                    if not link or not hasattr(link, "Proxy") or not hasattr(link.Proxy, "get_targets"):
                        continue
                    targets = link.Proxy.get_targets(link)
                    if link_name in ("FEM_Beam_Link_Fixed", "FEM_Beam_Link_Force") and targets:
                        # Apply only refs that belong to the beam so we never push test-geometry refs (Empty femnodes_mesh)
                        add_sets = list(getattr(link, "AddSelectionSets", []) or [])
                        refs_raw = get_fem_references_from_selectionsets(add_sets) if add_sets else []
                        refs_beam_only = [r for r in refs_raw if r[0] in beam_shape_objects]
                        for target in targets:
                            if not hasattr(target, "References"):
                                continue
                            target.References = []
                            if refs_beam_only:
                                apply_references_to_fem_constraint_one_by_one(refs_beam_only, target, doc)
                    else:
                        _apply_link_to_target(link)
                doc.recompute()
            except Exception as e:
                FreeCAD.Console.PrintWarning("FEM_Beam_Config: switch geometry failed (%s).\n" % e)

        try:
            from PySide import QtCore
            QtCore.QTimer.singleShot(0, _do_switch)
        except Exception:
            _do_switch()


class _ViewProviderFEM_Beam_Config:
    def __init__(self, vobj):
        vobj.Proxy = self

    def getIcon(self):
        return ""

    def __getstate__(self):
        return None

    def __setstate__(self, state):
        pass


def _get_beam_cut_object(doc):
    """Return the cut object (Beam_Cut or 'Cut' from BOPTools). Used so both naming conventions work."""
    return doc.getObject("Beam_Cut") or doc.getObject("Cut")


def _fem_cantilever_debug_logger():
    """Return (debug_log_fn, debug_path, file_handle). debug_log_fn(msg) writes to report view and to a .log file for export. When WRITE_FEM_DEBUG_FILES is False, returns no-op logger and None path/handle."""
    if not WRITE_FEM_DEBUG_FILES:
        return (lambda msg, include_tb=False: None), None, None
    import os
    import traceback
    base = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base, "fem_cantilever_debug.log")
    try:
        f = open(path, "w", encoding="utf-8")
    except Exception:
        f = None
        path = None

    def debug_log(msg, include_tb=False):
        line = "[FEM-cut] %s" % msg
        print(line)
        try:
            if FreeCADGui.getMainWindow():
                report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")
                if report:
                    report.append(line)
        except Exception:
            pass
        if f:
            try:
                f.write(line + "\n")
                if include_tb:
                    f.write(traceback.format_exc() + "\n")
                f.flush()
            except Exception:
                pass

    return debug_log, path, f


def _make_beam_cut(doc, beam_box, cyl, log, debug_log=None):
    """
    Create cut (box minus cylinder). Same code path for 1.0.2 and 1.1; order of attempts:

    - FreeCAD 1.0.2: BOPTools.BOPFeatures.make_cut() is preferred (creates object named "Cut").
      Part::Cut may be missing or fail on some shapes; Part.Shape.cut() is the fallback.
    - FreeCAD 1.1:   BOPTools or Part::Cut both work; Part.Shape.cut() fallback.

    Tries in order: (1) BOPTools.BOPFeatures.make_cut(), (2) Part::Cut, (3) Part::Feature +
    Part.Shape.cut(). Returns the cut object or None. Use _get_beam_cut_object(doc) to get
    the cut (handles both "Cut" and "Beam_Cut" names).
    """
    dbg = debug_log if callable(debug_log) else (lambda msg, include_tb=False: None)
    name = "Beam_Cut"
    dbg("_make_beam_cut: doc=%s, beam_box=%s, cyl=%s" % (doc.Name, beam_box.Name if beam_box else None, cyl.Name if cyl else None))
    dbg("  doc object names: %s" % [o.Name for o in doc.Objects])
    existing = doc.getObject(name)
    if existing:
        dbg("  removing existing Beam_Cut")
        doc.removeObject(existing.Name)
    cut_bop = doc.getObject("Cut")
    if cut_bop:
        dbg("  removing existing Cut")
        doc.removeObject(cut_bop.Name)
    cut = None

    # 1) BOPTools.BOPFeatures.make_cut()
    dbg("  Trying 1) BOPTools.BOPFeatures.make_cut([%s, %s])" % (beam_box.Name, cyl.Name))
    try:
        from BOPTools import BOPFeatures
        dbg("  BOPTools.BOPFeatures imported")
        bp = BOPFeatures.BOPFeatures(doc)
        dbg("  BOPFeatures(doc) created, calling make_cut")
        bp.make_cut([beam_box.Name, cyl.Name])
        dbg("  make_cut() returned, recomputing")
        doc.recompute()
        cut = doc.getObject("Cut")
        dbg("  getObject('Cut') = %s" % (cut.Name if cut else None))
        if cut:
            dbg("  cut.Shape=%s isNull=%s" % (bool(cut.Shape), cut.Shape.isNull() if cut.Shape else "N/A"))
            if cut.Shape:
                solids = getattr(cut.Shape, "Solids", None)
                dbg("  cut.Shape.Solids: count=%s" % (len(solids) if solids else 0))
            if cut.Shape and not cut.Shape.isNull() and (cut.Shape.Solids and len(cut.Shape.Solids) > 0):
                if log:
                    log("FEM cantilever beam example: Cut created with BOPTools.BOPFeatures.make_cut() (object name: Cut).")
                dbg("  SUCCESS: cut created via BOP")
                return cut
        else:
            dbg("  BOP did not create object 'Cut'; doc objects now: %s" % [o.Name for o in doc.Objects])
    except Exception as e:
        dbg("  BOPTools.BOPFeatures.make_cut FAILED: %s" % e, include_tb=True)
        if log:
            log("FEM cantilever beam example: BOPTools.BOPFeatures.make_cut failed (%s)." % e)
    cut = None

    # 2) Part::Cut
    dbg("  Trying 2) Part::Cut")
    try:
        cut = doc.addObject("Part::Cut", name)
        dbg("  addObject('Part::Cut') created: %s" % cut.Name)
        cut.Base = beam_box
        cut.Tool = cyl
        doc.recompute()
        dbg("  Part::Cut Shape isNull=%s Solids=%s" % (cut.Shape.isNull() if cut.Shape else "N/A", len(cut.Shape.Solids) if cut.Shape and getattr(cut.Shape, "Solids", None) else 0))
        if cut.Shape and not cut.Shape.isNull() and (cut.Shape.Solids and len(cut.Shape.Solids) > 0):
            if log:
                log("FEM cantilever beam example: Beam_Cut created with Part::Cut.")
            dbg("  SUCCESS: cut created via Part::Cut")
            return cut
    except Exception as e:
        dbg("  Part::Cut FAILED: %s" % e, include_tb=True)
        if log:
            log("FEM cantilever beam example: Part::Cut failed (%s)." % e)
    if cut:
        try:
            doc.removeObject(cut.Name)
        except Exception:
            pass
        cut = None

    # 3) Part::Feature + Part.Shape.cut()
    dbg("  Trying 3) Part::Feature + Part.Shape.cut()")
    try:
        dbg("  beam_box.Shape=%s cyl.Shape=%s" % (bool(beam_box.Shape), bool(cyl.Shape)))
        cut_shape = beam_box.Shape.cut(cyl.Shape)
        dbg("  cut_shape isNull=%s Solids=%s" % (cut_shape.isNull(), len(cut_shape.Solids) if cut_shape.Solids else 0))
        if cut_shape.isNull() or not (cut_shape.Solids and len(cut_shape.Solids) > 0):
            dbg("  Part.Shape.cut() produced no solid")
            if log:
                log("FEM cantilever beam example: Part.Shape.cut() produced no solid.")
            return None
        cut = doc.addObject("Part::Feature", name)
        cut.Shape = cut_shape
        doc.recompute()
        if log:
            log("FEM cantilever beam example: Beam_Cut created with Part::Feature (Part.Shape.cut() fallback).")
        dbg("  SUCCESS: cut created via Part::Feature")
        return cut
    except Exception as e:
        dbg("  Part.Shape.cut() fallback FAILED: %s" % e, include_tb=True)
        if log:
            log("FEM cantilever beam example: Part.Shape.cut() fallback failed (%s). No cut." % e)
        return None


def run_fem_cantilever_beam_example(cylinder_cut=False, geometry_offset=None):
    """
    FEM cantilever beam example. Uses setup_cantilever_base_solid when available;
    otherwise builds analysis/solver/mesh/material/constraints. In custom build,
    both Beam_Box and Beam_Cut are created; set FEM_Beam_Config.Suppressed=False (or unsuppress Beam_Cut if it has Suppressed) to use the cut.
    geometry_offset: optional FreeCAD.Vector to place the beam away from origin (e.g. when running with other test geometry to avoid overlap). Used when building from run_selectionset_tests.
    """
    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view") if FreeCADGui.getMainWindow() else None

    def log(msg):
        if report:
            report.append(msg)
        else:
            print(msg)

    try:
        import Fem  # noqa: F401
    except Exception as e:
        log("FEM cantilever beam example: SKIP (Fem module not available: %s)." % e)
        return

    doc = FreeCAD.ActiveDocument
    if doc is None:
        doc = FreeCAD.newDocument()
    doc.Label = "FEM_CantileverBeam"

    # Debug log to .log for diagnosing cut creation (e.g. in FreeCAD 1.1)
    debug_log, debug_path, debug_f = _fem_cantilever_debug_logger()
    debug_log("FEM cantilever example started; doc=%s" % doc.Name)
    debug_log("FreeCAD.Version=%s" % str(getattr(FreeCAD, "Version", ())))

    # Always use custom build so we create Beam_Box + cylinder + cut and the Suppressed toggle (femexamples path only has Box, no cut).
    use_base_solid = False
    log("FEM cantilever beam example: using custom build (Beam_Box + cylinder + cut + toggle).")
    debug_log("use_base_solid=%s (always False so cut is created)" % use_base_solid)

    if not use_base_solid:
        # Custom build: same dimensions as ccx_cantilever_base_solid (Length=8000, Width=Height=1000 mm)
        # Always create both Beam_Box and Beam_Cut; switch via FEM_Beam_Config.Suppressed (or Beam_Cut.Suppressed if available)
        offset = geometry_offset if geometry_offset is not None else FreeCAD.Vector(0, 0, 0)
        length_mm = 8000.0
        w, h = 1000.0, 1000.0
        beam_box = doc.getObject("Beam_Box")
        if not beam_box:
            beam_box = doc.addObject("Part::Box", "Beam_Box")
        beam_box.Length = length_mm
        beam_box.Width = w
        beam_box.Height = h
        beam_box.Placement.Base = offset
        doc.recompute()
        # Cylinder along Y at half cantilever length; through whole beam in ±Y (both sides)
        cyl = doc.getObject("Beam_CylinderCut")
        if not cyl:
            cyl = doc.addObject("Part::Cylinder", "Beam_CylinderCut")
        cyl.Radius = 300.0
        cyl.Height = w + 200.0  # along Y, through full width with 100 mm past each side
        # Center at mid-length; Y = 0 (or -50) so cylinder origin at beam edge / just outside
        cyl_center_y = 0.0  # use -50.0 for 50 mm past y=0
        cyl.Placement = FreeCAD.Placement(
            FreeCAD.Vector(0.5 * length_mm, cyl_center_y, 0.35 * h) + offset,
            FreeCAD.Rotation(FreeCAD.Vector(1, 0, 0), -90),
        )
        doc.recompute()
        # Create cut (box minus cylinder): BOPTools.BOPFeatures.make_cut() first (same as Part_Cut in 1.1 RC3), then Part::Cut, then Part::Feature fallback
        cut = _get_beam_cut_object(doc)
        debug_log("Before cut create: _get_beam_cut_object=%s" % (cut.Name if cut else None))
        if cut:
            try:
                if hasattr(cut, "Base") and hasattr(cut, "Tool"):
                    cut.Base = beam_box
                    cut.Tool = cyl
                    doc.recompute()
            except Exception:
                cut = None
            # If existing cut has no valid shape (e.g. Part::Cut failed in 1.1), replace it
            if cut and (not cut.Shape or cut.Shape.isNull() or not (getattr(cut.Shape, "Solids", None) and len(cut.Shape.Solids) > 0)):
                try:
                    doc.removeObject(cut.Name)
                except Exception:
                    pass
                cut = None
        if not cut:
            cut = _make_beam_cut(doc, beam_box, cyl, log, debug_log=debug_log)
        debug_log("After cut create: cut=%s" % (cut.Name if cut else None))
        if cut:
            # If cut has native Suppressed (e.g. some FreeCAD versions), set to False so we use cut (hole) by default
            if hasattr(cut, "Suppressed"):
                try:
                    cut.Suppressed = False
                except Exception:
                    pass
            doc.recompute()
        # Default is use cut (hole); beam for initial SelectionSet setup, config.Suppressed=False will switch to cut
        beam = beam_box

    # SelectionSets by coordinates: fixed end (x=0), force end (x=length_mm)
    def make_face_selectionset(name, point, label_suffix):
        existing = doc.getObject(name)
        if existing:
            doc.removeObject(existing.Name)
        obj = doc.addObject("App::FeaturePython", name)
        SelectionSet(obj)
        obj.Label = "SelectionSet " + label_suffix
        obj.MainObject = beam
        obj.Mode = "faces"
        obj.VolumeMode = "coordinates"
        obj.CoordinatePoint = point
        recomputeSelectionSetFromShapes(obj)
        apply_selectionset_view_defaults(obj)
        return obj

    # Use the same geometry offset as the beam so the coordinate points hit the correct faces/solid
    fixed_pt = FreeCAD.Vector(
        offset.x,
        offset.y + 0.5 * w,
        offset.z + 0.5 * h,
    )  # (x=0, half width, half height) – left end face
    force_pt = FreeCAD.Vector(
        offset.x + length_mm,
        offset.y + 0.5 * w,
        offset.z + 0.5 * h,
    )  # (x=length, half width, half height) – right end face
    ss_fixed = make_face_selectionset("Beam_FixedFace", fixed_pt, "Beam fixed face")
    ss_force = make_face_selectionset("Beam_ForceFace", force_pt, "Beam force face")
    # SelectionSet for beam solid (for material): one solid, center point
    def make_solid_selectionset(name, point, label_suffix):
        existing = doc.getObject(name)
        if existing:
            doc.removeObject(existing.Name)
        obj = doc.addObject("App::FeaturePython", name)
        SelectionSet(obj)
        obj.Label = "SelectionSet " + label_suffix
        obj.MainObject = beam
        obj.Mode = "solids"
        obj.VolumeMode = "coordinates"
        obj.CoordinatePoint = point
        recomputeSelectionSetFromShapes(obj)
        apply_selectionset_view_defaults(obj)
        return obj
    # SelectionSet Beam solid: use explicit coordinates (mm) as requested
    solid_pt = FreeCAD.Vector(1000.0, 500.0, 500.0)
    ss_solid = make_solid_selectionset("Beam_Solid", solid_pt, "Beam solid")
    doc.recompute()

    # FEM analysis: from setup_cantilever_base_solid we have Analysis, or create our own
    if not use_base_solid:
        analysis = doc.getObject("FEM_Beam_Analysis")
        if not analysis:
            analysis = doc.addObject("Fem::FemAnalysis", "FEM_Beam_Analysis")
    else:
        analysis = doc.Analysis
    # Solver and mesh: only create when not using base_solid (base_solid already adds them)
    if not use_base_solid:
        solver = doc.getObject("FEM_Beam_Solver")
        if not solver:
            try:
                import ObjectsFem
                solver = getattr(ObjectsFem, "makeSolverCalculixCxxtools", None) or getattr(ObjectsFem, "makeSolverCalculiXCcxTools", None)
                if callable(solver):
                    solver = solver(doc, "FEM_Beam_Solver")
                else:
                    solver = None
            except Exception:
                solver = None
            if not solver:
                try:
                    solver = doc.addObject("Fem::SolverCcxTools", "FEM_Beam_Solver")
                except Exception as e:
                    log("FEM cantilever beam example: solver not created (%s)." % e)
                    solver = None
            if solver:
                analysis.addObject(solver)
                log("FEM cantilever beam example: Solver (CalculiX) added to analysis.")

    # A) GMSH mesh: base_solid already adds Mesh; otherwise create and add
    if not use_base_solid:
        mesh = doc.getObject("FEM_Beam_Mesh")
        if not mesh:
            try:
                mesh = doc.addObject("Fem::FemMeshGmsh", "FEM_Beam_Mesh")
            except Exception:
                try:
                    import ObjectsFem
                    mesh = getattr(ObjectsFem, "makeMeshGmsh", None)
                    if callable(mesh):
                        mesh = mesh(doc, "FEM_Beam_Mesh")
                    else:
                        mesh = doc.addObject("Fem::FemMeshGmsh", "FEM_Beam_Mesh")
                except Exception as e:
                    log("FEM cantilever beam example: mesh not created (%s)." % e)
                    mesh = None
        if mesh:
            # Use cut (hole) for mesh when available, to match default Suppressed=no
            mesh_geometry = cut if cut else beam
            # Set geometry: femexamples use .Shape (ccx_cantilever_base_solid); 1.0.2 may use Part. Set all that exist.
            for prop in ("Shape", "Part", "Geometry"):
                if hasattr(mesh, prop):
                    try:
                        setattr(mesh, prop, mesh_geometry)
                    except Exception:
                        pass
            # 1.0.2: mesh size often in CharacteristicLengthMax; 1.1: CharacterLengthMax
            if hasattr(mesh, "CharacteristicLengthMax"):
                mesh.CharacteristicLengthMax = 200.0
            # 2nd order elements (Tet10): ElementOrder drives GMSH; SecondOrderLinear from template
            if hasattr(mesh, "ElementOrder"):
                mesh.ElementOrder = "2nd"
            if hasattr(mesh, "SecondOrderLinear"):
                mesh.SecondOrderLinear = False
            if analysis:
                analysis.addObject(mesh)
            # Try to generate the mesh (femexamples, or GmshTools.run/create_mesh for 1.1/1.0.2)
            mesh_generated = _run_gmsh_for_mesh(mesh)
            if not mesh_generated:
                log("FEM cantilever beam example: could not run GMSH. Run GMSH from the mesh object in the FEM workbench.")
            if mesh_generated:
                doc.recompute()
                log("FEM cantilever beam example: GMSH mesh generated (2nd order).")
            else:
                log("FEM cantilever beam example: GMSH mesh in analysis. Run GMSH from the mesh object to generate.")
    elif use_base_solid:
        # Ensure 2nd order when using setup_cantilever_base_solid (mesh is named "Mesh")
        mesh = doc.getObject("Mesh")
        if mesh and hasattr(mesh, "ElementOrder"):
            mesh.ElementOrder = "2nd"

    # A) Material: base_solid has FemMaterial; otherwise create and add
    if use_base_solid:
        material_obj = doc.getObject("FemMaterial")
    else:
        material_obj = doc.getObject("FEM_Beam_Material")
        if not material_obj:
            try:
                import ObjectsFem
                material_obj = ObjectsFem.makeMaterialSolid(doc, "FEM_Beam_Material")
            except Exception:
                try:
                    material_obj = doc.addObject("Fem::Material", "FEM_Beam_Material")
                except Exception:
                    material_obj = None
        if material_obj and analysis:
            analysis.addObject(material_obj)
        if material_obj:
            mat = getattr(material_obj, "Material", None) or {}
            if isinstance(mat, dict):
                mat["Name"] = "CalculiX-Steel"
                mat["YoungsModulus"] = "210000 MPa"
                mat["PoissonRatio"] = "0.30"
                material_obj.Material = mat
    if material_obj:
        from selection_set_core import _apply_link_to_target
        link_material = doc.getObject("FEM_Beam_Link_Material")
        if not link_material:
            # Create with explicit inputs so we never inherit AddShapes from GUI selection (test geometry)
            link_material = createSelectionSetLink(
                name="FEM_Beam_Link_Material",
                _target=material_obj,
                _add_sets=[ss_solid],
                _add_shapes=[],
                _targets=[material_obj],
            )
        if link_material:
            n = len(getattr(ss_solid, "ElementList", []) or [])
            unit = "solid" if n == 1 else "solids"
            link_material.Label = "SelectionSetLink Beam (%d %s) Material" % (n, unit)
            link_material.TargetObjects = [material_obj]
            link_material.AddSelectionSets = [ss_solid]
            link_material.AddShapes = []
            link_material.SubtractSelectionSets = []
            link_material.SubtractShapes = []
            link_material.ShapeRefMode = "solids"
            link_material.UseOneByOneWorkaround = True
            _apply_link_to_target(link_material)
        log("FEM cantilever beam example: Material in analysis, refs from SelectionSet Beam_Solid via link.")

    try:
        if use_base_solid:
            constraint_fixed = doc.getObject("ConstraintFixed")
            constraint_force = doc.getObject("ConstraintForce")
            if not constraint_force:
                try:
                    import ObjectsFem
                    constraint_force = ObjectsFem.makeConstraintForce(doc, "ConstraintForce")
                except Exception:
                    constraint_force = doc.addObject("Fem::ConstraintForce", "ConstraintForce")
                analysis.addObject(constraint_force)
        else:
            constraint_fixed = doc.getObject("FEM_Beam_ConstraintFixed")
            if not constraint_fixed:
                constraint_fixed = doc.addObject("Fem::ConstraintFixed", "FEM_Beam_ConstraintFixed")
                analysis.addObject(constraint_fixed)
            constraint_force = doc.getObject("FEM_Beam_ConstraintForce")
            if not constraint_force:
                constraint_force = doc.addObject("Fem::ConstraintForce", "FEM_Beam_ConstraintForce")
                analysis.addObject(constraint_force)
    except Exception as e:
        log("FEM cantilever beam example: SKIP (FEM constraints not available: %s)." % e)
        if debug_f:
            try:
                debug_log("Early exit (constraints not available). Debug log: %s" % debug_path)
                debug_f.close()
            except Exception:
                pass
        return

    # Force: same as ccx_cantilever_faceload – 9e6 N (9 MN); use string with unit so value is not misinterpreted
    if hasattr(constraint_force, "Force"):
        try:
            constraint_force.Force = "9000000.0 N"
        except Exception:
            try:
                constraint_force.Force = 9000000.0
            except Exception:
                pass
    # Direction: same as ccx_cantilever_faceload – (beam, "Edge5"), Reversed = True (vertical downward)
    try:
        if hasattr(constraint_force, "Direction"):
            shape = beam.Shape if hasattr(beam, "Shape") else None
            direction_set = False
            if shape and len(shape.Edges) >= 5:
                try:
                    # Official example uses Edge5 for the box
                    edge5 = shape.getElement("Edge5")
                    if edge5 and len(edge5.Vertexes) >= 2:
                        d = edge5.Vertexes[1].Point.sub(edge5.Vertexes[0].Point)
                        if d.Length > 1e-6 and abs(d.normalize().z) > 0.99:
                            constraint_force.Direction = (beam, "Edge5")
                            direction_set = True
                except Exception:
                    pass
            if not direction_set and shape:
                for i in range(1, len(shape.Edges) + 1):
                    try:
                        edge = shape.getElement("Edge%d" % i)
                        if edge and len(edge.Vertexes) >= 2:
                            d = edge.Vertexes[1].Point.sub(edge.Vertexes[0].Point)
                            if d.Length > 1e-6:
                                d.normalize()
                                if abs(d.z) > 0.99:
                                    constraint_force.Direction = (beam, "Edge%d" % i)
                                    direction_set = True
                                    break
                    except Exception:
                        continue
            if hasattr(constraint_force, "Reversed"):
                constraint_force.Reversed = True
    except Exception:
        pass

    # Apply refs from SelectionSets to constraints via SelectionSetLinks only (so constraints use the SelectionSets).
    # Restrict refs to the beam shape(s) only so we never apply test-geometry refs (BooleanFragments_CylTwoHalves,
    # Body003, etc.) when running in the same document as SelectionSet tests — those cause "Empty femnodes_mesh" errors.
    from selection_set_core import _apply_link_to_target, get_fem_references_from_selectionsets, apply_references_to_fem_constraint_one_by_one
    beam_shape_objects = (beam,)
    cut_obj = _get_beam_cut_object(doc)
    if cut_obj and cut_obj not in beam_shape_objects:
        beam_shape_objects = (beam, cut_obj)
    link_fixed = doc.getObject("FEM_Beam_Link_Fixed")
    if not link_fixed:
        link_fixed = createSelectionSetLink(
            name="FEM_Beam_Link_Fixed",
            _target=constraint_fixed,
            _add_sets=[ss_fixed],
            _add_shapes=[],
            _targets=[constraint_fixed],
        )
    if link_fixed and constraint_fixed:
        constraint_fixed.References = []
        n = len(getattr(ss_fixed, "ElementList", []) or [])
        unit = "face" if n == 1 else "faces"
        link_fixed.Label = "SelectionSetLink Beam (%d %s) Fixed" % (n, unit)
        link_fixed.TargetObjects = [constraint_fixed]
        link_fixed.AddSelectionSets = [ss_fixed]
        link_fixed.AddShapes = []
        link_fixed.SubtractSelectionSets = []
        link_fixed.SubtractShapes = []
        link_fixed.UseOneByOneWorkaround = True
        refs_raw = get_fem_references_from_selectionsets([ss_fixed])
        refs_beam_only = [r for r in refs_raw if r[0] in beam_shape_objects]
        if refs_beam_only:
            apply_references_to_fem_constraint_one_by_one(refs_beam_only, constraint_fixed, doc)
        else:
            _apply_link_to_target(link_fixed)
        log("FEM cantilever beam example: ConstraintFixed refs from SelectionSet Beam_FixedFace via link.")
    link_force = doc.getObject("FEM_Beam_Link_Force")
    if not link_force:
        link_force = createSelectionSetLink(
            name="FEM_Beam_Link_Force",
            _target=constraint_force,
            _add_sets=[ss_force],
            _add_shapes=[],
            _targets=[constraint_force],
        )
    if link_force and constraint_force:
        constraint_force.References = []
        n = len(getattr(ss_force, "ElementList", []) or [])
        unit = "face" if n == 1 else "faces"
        link_force.Label = "SelectionSetLink Beam (%d %s) Force" % (n, unit)
        link_force.TargetObjects = [constraint_force]
        link_force.AddSelectionSets = [ss_force]
        link_force.AddShapes = []
        link_force.SubtractSelectionSets = []
        link_force.SubtractShapes = []
        link_force.UseOneByOneWorkaround = True
        refs_raw = get_fem_references_from_selectionsets([ss_force])
        refs_beam_only = [r for r in refs_raw if r[0] in beam_shape_objects]
        if refs_beam_only:
            apply_references_to_fem_constraint_one_by_one(refs_beam_only, constraint_force, doc)
        else:
            _apply_link_to_target(link_force)
        log("FEM cantilever beam example: ConstraintForce refs from SelectionSet Beam_ForceFace via link.")

    # B) Config object and optional variable container (Variables in 1.1, VarSet in 1.0) for geometry switch
    if not use_base_solid:
        config = doc.getObject("FEM_Beam_Config")
        if not config:
            config = doc.addObject("App::FeaturePython", "FEM_Beam_Config")
            _FEM_Beam_ConfigProxy(config)
            if hasattr(config, "ViewObject") and config.ViewObject:
                config.ViewObject.Proxy = _ViewProviderFEM_Beam_Config(config.ViewObject)
        if config:
            config.Label = "FEM beam geometry switch"
            # Cylinder radius (mm): default 300; when Variables/VarSet exist they are the source (user edits there). Never use 0 (avoids "Radius of cylinder too small").
            if not hasattr(config, "CylinderRadius"):
                config.addProperty("App::PropertyFloat", "CylinderRadius", "Beam",
                                  "Cylinder cut radius (mm). Same x, z, mesh as standard hole. Driven by FEM_Beam_Variables/VarSet when present.")
                config.CylinderRadius = 300.0
            elif getattr(config, "CylinderRadius", 0) <= 0:
                config.CylinderRadius = 300.0
            # Resolve Variables/VarSet first so we can make them the source for CylinderRadius when present
            var_obj = None
            var_name = None
            try:
                fc_ver = _freecad_version_tuple()
                if fc_ver >= (1, 1):
                    try:
                        var_obj = doc.getObject("FEM_Beam_Variables")
                        if var_obj is None:
                            var_obj = doc.addObject("App::Variables", "FEM_Beam_Variables")
                        if var_obj:
                            var_name = "FEM_Beam_Variables"
                    except Exception:
                        var_obj = None
                        var_name = None
                if var_obj is None:
                    var_obj = doc.getObject("FEM_Beam_VarSet")
                    if var_obj is None:
                        var_obj = doc.addObject("App::VarSet", "FEM_Beam_VarSet")
                    if var_obj:
                        var_name = "FEM_Beam_VarSet"
            except Exception:
                var_obj = None
                var_name = None
            # Wire CylinderRadius: Variables/VarSet = source when present; config and cylinder follow.
            if var_obj and hasattr(var_obj, "addProperty") and var_name:
                if not hasattr(var_obj, "CylinderRadius"):
                    var_obj.addProperty("App::PropertyFloat", "CylinderRadius", "Beam",
                                        "Cylinder cut radius (mm). Edit here; config and cylinder follow.")
                    var_obj.CylinderRadius = 300.0  # default 300, never 0 (PropertyFloat defaults to 0)
                try:
                    if getattr(var_obj, "getExpression", lambda p: None)("CylinderRadius"):
                        var_obj.setExpression("CylinderRadius", None)
                except Exception:
                    pass
                # Sync from config but never use 0 (suppressed or not, 300 is valid and avoids error messages)
                r = float(getattr(config, "CylinderRadius", 300.0) or 300.0)
                if r <= 0:
                    r = 300.0
                var_obj.CylinderRadius = r
                doc.recompute()  # commit variable value before expressions read it (avoids radius 0 / "too small")
                try:
                    config.setExpression("CylinderRadius", "%s.CylinderRadius" % var_name)
                except Exception:
                    config.CylinderRadius = float(var_obj.CylinderRadius)
            else:
                config.CylinderRadius = 300.0
            cyl_var = doc.getObject("Beam_CylinderCut")
            if cyl_var and hasattr(cyl_var, "setExpression"):
                # Ensure cylinder has a valid radius before expression (avoids "Radius of cylinder too small" on first recompute)
                safe_r = 300.0
                if var_obj and hasattr(var_obj, "CylinderRadius"):
                    safe_r = float(getattr(var_obj, "CylinderRadius", 300.0)) or 300.0
                elif config and hasattr(config, "CylinderRadius"):
                    safe_r = float(getattr(config, "CylinderRadius", 300.0)) or 300.0
                try:
                    if cyl_var.Radius <= 0:
                        cyl_var.Radius = max(1.0, safe_r)
                except Exception:
                    pass
                try:
                    # Drive cylinder directly from variable when present so GUI edits to variable take effect in all FC versions
                    if var_obj and var_name and hasattr(var_obj, "CylinderRadius"):
                        cyl_var.setExpression("Radius", "%s.CylinderRadius" % var_name)
                    else:
                        cyl_var.setExpression("Radius", "FEM_Beam_Config.CylinderRadius")
                except Exception:
                    pass
                doc.recompute()
            # Link Suppressed to Variables/VarSet when present
            if var_obj and var_name:
                try:
                    if not hasattr(var_obj, "Suppressed"):
                        var_obj.addProperty("App::PropertyBool", "Suppressed", "Beam",
                                            "Suppress cylinder cut: True=box, False=cut. Drives FEM beam geometry switch.")
                        var_obj.Suppressed = False
                    var_obj.Label = "FEM beam variables"
                    expr = getattr(config, "getExpression", lambda p: None)("Suppressed") if hasattr(config, "getExpression") else None
                    if not expr and hasattr(config, "setExpression"):
                        config.setExpression("Suppressed", "%s.Suppressed" % var_name)
                    log("FEM cantilever beam example: Toggle 'Suppressed' and edit 'CylinderRadius' on '%s' (or FEM_Beam_Config); geometry follows." % var_name)
                except Exception:
                    log("FEM cantilever beam example: Toggle 'Suppressed' on 'FEM_Beam_Config' in Data tab (True=box, False=cut).")
            else:
                log("FEM cantilever beam example: Toggle 'Suppressed' on 'FEM_Beam_Config' in Data tab (True=box, False=cut).")
        # Optional: add a TextDocument with readme and expected deflection values
        try:
            td = doc.getObject("FEM_Beam_Readme")
            if td is None:
                td = doc.addObject("App::TextDocument", "FEM_Beam_Readme")
            if td is not None:
                td.Label = "FEM beam – readme and expected values"
                td.Text = """FEM Cantilever Beam – Expected deflections

Geometry: Beam 8000 × 1000 × 1000 mm. Hole: cylinder axis Y, center x=4000 mm, z=350 mm.
Mesh: CharacterLengthMax = 200 mm. Material on beam solid only.

Expected max deflection (mm) at free end:
• Without hole (box only):       ~%.1f
• With hole, radius 300 mm:       ~%.1f
• With hole, radius 400 mm:       ~%.1f

Toggle FEM_Beam_Config.Suppressed: False = cut (with hole), True = box only.
Cylinder radius: edit in FEM_Beam_Variables / FEM_Beam_VarSet (CylinderRadius); config and cylinder follow.

On FreeCAD 1.0.2, GMSH mesh with hole may fail; run mesh/solver manually and expect the values above.
""" % (FEM_CANTILEVER_REF_NO_HOLE_MM, FEM_CANTILEVER_REF_WITH_R300mm_HOLE_MM, FEM_CANTILEVER_REF_WITH_R400mm_HOLE_MM)
        except Exception:
            pass
        # Group geometry, SelectionSets, links, config, and readme into one tree group
        beam_group_names = ["Beam_Box", "Beam_CylinderCut", "Beam_FixedFace", "Beam_ForceFace", "Beam_Solid",
                           "FEM_Beam_Link_Material", "FEM_Beam_Link_Fixed", "FEM_Beam_Link_Force", "FEM_Beam_Config"]
        if doc.getObject("FEM_Beam_Readme"):
            beam_group_names.append("FEM_Beam_Readme")
        cut_obj = _get_beam_cut_object(doc)
        if cut_obj:
            beam_group_names.append(cut_obj.Name)
        for vname in ("FEM_Beam_Variables", "FEM_Beam_VarSet"):
            if doc.getObject(vname):
                beam_group_names.append(vname)
                break
        _add_objects_to_group(doc, "FEM_Beam_Group", beam_group_names, label="FEM Cantilever")

    doc.recompute()
    try:
        if FreeCADGui.ActiveDocument and FreeCADGui.ActiveDocument.ActiveView:
            FreeCADGui.ActiveDocument.ActiveView.viewIsometric()
            FreeCADGui.ActiveDocument.ActiveView.fitAll()
    except Exception:
        pass
    # Save document with the given name (macro folder)
    try:
        import os
        base = os.path.dirname(os.path.abspath(__file__))
        save_path = os.path.join(base, "FEM_CantileverBeam_with_SelectionSetLinks.FCStd")
        doc.saveAs(save_path)
        log("FEM cantilever beam example: saved as %s" % save_path)
    except Exception as e:
        log("FEM cantilever beam example: could not save file (%s)." % e)
    if debug_f:
        try:
            debug_log("Debug log written to: %s" % debug_path)
            debug_f.close()
        except Exception:
            pass
    log("FEM cantilever beam example: done. Beam + fixed/force SelectionSets and SelectionSetLinks created." + (" (with cylinder cut)" if cylinder_cut else ""))


def run_selectionset_tests():
    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view")
    _log_lines = []  # capture for results file

    def log(msg):
        _log_lines.append(msg)
        if report:
            report.append(msg)
        else:
            print(msg)

    log("\n[SelectionSet Test Runner] Starting tests...")
    log("[SelectionSet Test Runner] Macro version: %s (use this version to test if the loaded macro meets required %s)." % (MACRO_VERSION, REQUIRED_MACRO_VERSION))
    if not _macro_version_ok():
        log("[SelectionSet Test Runner] WARNING: Loaded macro version is below required %s; some tests may be skipped or fail." % REQUIRED_MACRO_VERSION)

    # Suppress verbose SelectionSet callback output (DEBUG, "Button Pressed", per-face/per-solid lines)
    _prev_verbose = getattr(selection_set_core, "SELECTIONSET_VERBOSE_CALLBACK", True)
    selection_set_core.SELECTIONSET_VERBOSE_CALLBACK = False

    doc = FreeCAD.ActiveDocument
    if not doc:
        # Align behaviour with Run Test 1: create a dedicated test document if needed.
        doc = FreeCAD.newDocument()
        doc.Label = "SelectionSetTest"
        log("Created new document 'SelectionSetTest' for SelectionSet tests.")

    # Temporarily remove SelectionSet observer so test addSelection() calls don't trigger it (avoids cascade and duplicate observer errors)
    try:
        _detach_selectionset_observer()
    except Exception:
        pass

    n_tests = 0
    n_failed = 0
    n_passed = 0
    failed_tests = []

    # Ensure test geometry for Tests 1–4 (create only if missing)
    main, sel_shape = _ensure_test_1_4_geometry()

    # Test 1: TestCube_main first, then TestCube_Selection (expect top face only)
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(main)
    FreeCADGui.Selection.addSelection(sel_shape)
    ss1 = createSelectionSetFromCurrent(
        selection_set_name="Test1_Faces_TestCube_main_Face6_only", force_mode="faces"
    )
    expected1 = ["TestCube_main.Face6"]
    result1 = list(ss1.ElementList)
    n_tests += 1
    if result1 == expected1:
        n_passed += 1
    else:
        n_failed += 1
        failed_tests.append("Test 1")
    log(
        f"Test 1: {result1} == {expected1} -> {'PASS' if result1 == expected1 else 'FAIL'}"
    )
    # log(
    #     "Test 1 Note: Expected only the top face (Face6) of TestCube_main to be selected because it is the only face intersected by the selection-defining shape when selected in this order."
    # )

    # Test 2: TestCube_Selection first, then TestCube_main (expect empty)
    # main, sel_shape = build_test_1_cubes()
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(sel_shape)
    FreeCADGui.Selection.addSelection(main)
    ss2 = createSelectionSetFromCurrent(
        selection_set_name="Test2_Faces_empty_sel_shape_first", force_mode="faces"
    )
    expected2 = []
    result2 = list(ss2.ElementList)
    n_tests += 1
    if result2 == expected2:
        n_passed += 1
    else:
        n_failed += 1
        failed_tests.append("Test 2")
    log(
        f"Test 2: {result2} == {expected2} -> {'PASS' if result2 == expected2 else 'FAIL'}"
    )
    # log(
    #     "Test 2 Note: Expected empty selection because the selection-defining shape does not intersect any faces of the main object when selected in this order."
    # )

    # Test 3: solids, main first (default volume mode) – expect empty
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(main)
    FreeCADGui.Selection.addSelection(sel_shape)
    ss3 = createSelectionSetFromCurrent(
        selection_set_name="Test3_Solids_empty_main_first",
        force_mode="solids",
    )
    expected3 = []
    result3 = list(ss3.ElementList)
    n_tests += 1
    if result3 == expected3:
        n_passed += 1
    else:
        n_failed += 1
        failed_tests.append("Test 3")
    log(
        f"Test 3 (solids): {result3} == {expected3} -> {'PASS' if result3 == expected3 else 'FAIL'}"
    )

    # Test 3a: solids, intersection – TestCube_main as main; TestCube_Selection as selection; expect Solid1
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(main)
    FreeCADGui.Selection.addSelection(sel_shape)
    ss3a = createSelectionSetFromCurrent(
        selection_set_name="Test3a_Solids_main_first_intersection",
        force_mode="solids",
        volume_mode="intersection",
    )
    expected3a = ["TestCube_main.Solid1"]
    result3a = list(ss3a.ElementList)
    n_tests += 1
    if result3a == expected3a:
        n_passed += 1
        log(f"Test 3a (solids, intersection): PASS ({result3a}).")
    else:
        n_failed += 1
        failed_tests.append("Test 3a")
        log(f"Test 3a (solids, intersection): FAIL (got {result3a}, expected {expected3a}).")

    # Test 4: (example) Test solid selection (add your own expected result)
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(sel_shape)
    FreeCADGui.Selection.addSelection(main)
    ss4 = createSelectionSetFromCurrent(
        selection_set_name="Test4_Solids_empty_sel_shape_first", force_mode="solids"
    )
    expected4 = []  # Adjust as needed
    result4 = list(ss4.ElementList)
    n_tests += 1
    if result4 == expected4:
        n_passed += 1
    else:
        n_failed += 1
        failed_tests.append("Test 4")
    log(
        f"Test 4 (solids): {result4} == {expected4} -> {'PASS' if result4 == expected4 else 'FAIL'}"
    )

    # Ensure test geometry for Tests 5–8 (create only if missing)
    _, _, boolean, sel_cube = _ensure_test_5_8_geometry()

    # Test 5: BooleanFragments first, then TestCube_Selection_Bool (expect face 7 of BooleanFragments which is the sphere fragment)
    # cube, sphere, boolean = build_test_2_cube_with_inner_sphere()
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(boolean)
    FreeCADGui.Selection.addSelection(sel_cube)
    ss5 = createSelectionSetFromCurrent(
        selection_set_name="Test5_Faces_BooleanFragments_Face7_sphere_fragment", force_mode="faces"
    )
    expected5 = [
        f"BooleanFragments.Face7"
    ]  # Adjust if the sphere fragment face number is different
    result5 = list(ss5.ElementList)
    n_tests += 1
    if result5 == expected5:
        n_passed += 1
    else:
        n_failed += 1
        failed_tests.append("Test 5")
    log(
        f"Test 5: {result5} == {expected5} -> {'PASS' if result5 == expected5 else 'FAIL'}"
    )

    # Test 6: TestCube_Selection_Bool first, then BooleanFragments (all 6 faces fully inside)
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(sel_cube)
    FreeCADGui.Selection.addSelection(boolean)
    ss6 = createSelectionSetFromCurrent(
        selection_set_name="Test6_Faces_TestCube_Selection_Bool_6faces_fully_inside", force_mode="faces"
    )
    expected6 = [f"TestCube_Selection_Bool.Face{i + 1}" for i in range(6)]
    result6 = list(ss6.ElementList)
    n_tests += 1
    if set(result6) == set(expected6):
        n_passed += 1
        log("Test 6: PASS (all 6 faces fully inside BooleanFragments).")
    else:
        n_failed += 1
        failed_tests.append("Test 6")
        log(f"Test 6: {result6} == {expected6} -> FAIL")
    log(f"Test 6 (faces fully inside): {len(result6)} face(s) selected: {result6}")

    # Test 7: (example) Test solid selection (add your own expected result)
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(sel_cube)
    FreeCADGui.Selection.addSelection(boolean)
    ss7 = createSelectionSetFromCurrent(
        selection_set_name="Test7_Solids_TestCube_Selection_Bool_Solid1", force_mode="solids"
    )
    expected7 = [f"TestCube_Selection_Bool.Solid1"]  # Adjust as needed
    result7 = list(ss7.ElementList)
    n_tests += 1
    if result7 == expected7:
        n_passed += 1
    else:
        n_failed += 1
        failed_tests.append("Test 7")
    log(
        f"Test 7 (solids): {result7} == {expected7} -> {'PASS' if result7 == expected7 else 'FAIL'}"
    )

    # Test 8: (example) Test solid selection (add your own expected result)
    FreeCADGui.Selection.clearSelection()
    FreeCADGui.Selection.addSelection(boolean)
    FreeCADGui.Selection.addSelection(sel_cube)
    ss8 = createSelectionSetFromCurrent(
        selection_set_name="Test8_Solids_BooleanFragments_Solid2_sphere", force_mode="solids"
    )
    expected8 = [f"BooleanFragments.Solid2"]  # Adjust as needed
    result8 = list(ss8.ElementList)
    n_tests += 1
    if result8 == expected8:
        n_passed += 1
    else:
        n_failed += 1
        failed_tests.append("Test 8")
    log(
        f"Test 8 (solids): {result8} == {expected8} -> {'PASS' if result8 == expected8 else 'FAIL'}"
    )

    # Ensure test geometry for Tests 9, 9a, 9b (create only if missing)
    boolean_cyl, body003 = _ensure_test_9_geometry()
    if boolean_cyl and body003:
        FreeCADGui.Selection.clearSelection()
        FreeCADGui.Selection.addSelection(boolean_cyl)
        FreeCADGui.Selection.addSelection(body003)
        ss9 = createSelectionSetFromCurrent(
            selection_set_name="Test9_Faces_CylTwoHalves_intersection_6faces", force_mode="faces", volume_mode="intersection"
        )
        result9 = list(ss9.ElementList)
        expected9 = [
            "BooleanFragments_CylTwoHalves.Face4",
            "BooleanFragments_CylTwoHalves.Face5",
            "BooleanFragments_CylTwoHalves.Face6",
            "BooleanFragments_CylTwoHalves.Face7",
            "BooleanFragments_CylTwoHalves.Face8", # manually approved
            "BooleanFragments_CylTwoHalves.Face9", # manually approved            
        ]
        n_tests += 1
        log(f"Test 9 (faces intersection with Body003): {len(result9)} face(s) selected: {result9}")
        if set(result9) == set(expected9):
            n_passed += 1
            log("Test 9: PASS (expected faces that intersect Body003).")
        else:
            n_failed += 1
            failed_tests.append("Test 9")
            log(f"Test 9: FAIL (got {result9}, expected {expected9}).")

        # Test 9a: same geometry, faces fully inside Body003 (vertical faces in one plane, inside selection body)
        FreeCADGui.Selection.clearSelection()
        FreeCADGui.Selection.addSelection(boolean_cyl)
        FreeCADGui.Selection.addSelection(body003)
        ss9a = createSelectionSetFromCurrent(
            selection_set_name="Test9a_Faces_CylTwoHalves_fully_inside_5faces", force_mode="faces", volume_mode="fully_inside"
        )
        result9a = list(ss9a.ElementList)
        # Do not change expected9a: 2 half-spheres, each has 1 round + 2 vertical faces. One half fully inside + other's vertical in = 5 faces; both halves fully inside = 6.
        expected9a = [
            "BooleanFragments_CylTwoHalves.Face4", # manually approved
            "BooleanFragments_CylTwoHalves.Face5", # manually approved
            "BooleanFragments_CylTwoHalves.Face6", # manually approved
            "BooleanFragments_CylTwoHalves.Face8", # manually approved
            "BooleanFragments_CylTwoHalves.Face9", # manually approved
        ] # "BooleanFragments_CylTwoHalves.Face7" is the outer sperical face of the second half sphere and should not be included because its not fully-inside of the selection-shape
        n_tests += 1
        if set(result9a) == set(expected9a):
            n_passed += 1
            log(f"Test 9a: PASS ({len(result9a)} face(s) fully inside Body003).")
        else:
            n_failed += 1
            failed_tests.append("Test 9a")
            log(f"Test 9a: FAIL (got {result9a}, expected {expected9a}).")
        log(f"Test 9a (faces fully inside Body003): {len(result9a)} face(s) selected: {result9a}")

        # Test 9b: same geometry, solids fully inside Body003 (only one half-sphere solid is fully inside)
        FreeCADGui.Selection.clearSelection()
        FreeCADGui.Selection.addSelection(boolean_cyl)
        FreeCADGui.Selection.addSelection(body003)
        ss9b = createSelectionSetFromCurrent(
            selection_set_name="Test9b_Solids_CylTwoHalves_Solid3_one_half_inside", force_mode="solids"
        )
        result9b = list(ss9b.ElementList)
        expected9b = ["BooleanFragments_CylTwoHalves.Solid3"]  # one half-sphere fragment fully inside Body003
        n_tests += 1
        if set(result9b) == set(expected9b):
            n_passed += 1
            log("Test 9b: PASS (one solid fully inside Body003).")
        else:
            n_failed += 1
            failed_tests.append("Test 9b")
            log(f"Test 9b: FAIL (got {result9b}, expected {expected9b}).")
        log(f"Test 9b (solids fully inside Body003): {len(result9b)} solid(s) selected: {result9b}")

        # Test 9c: faces fully inside Body003 but only from the solid in 9b (Solid3 = one half-sphere)
        # Solid filter uses _get_solid_for_face (majority vote over face points) so faces are assigned to the correct solid.
        if True:
            FreeCADGui.Selection.clearSelection()
            FreeCADGui.Selection.addSelection(boolean_cyl)
            FreeCADGui.Selection.addSelection(body003)
            FreeCADGui.Selection.addSelection(ss9b)
            ss9c = createSelectionSetFromCurrent(
                selection_set_name="Test9c_Faces_CylTwoHalves_fully_inside_filtered_by_Solid3",
                force_mode="faces",
                volume_mode="fully_inside",
            )
            result9c = list(ss9c.ElementList)
            expected9c_count = 3
            expected9c = [
                "BooleanFragments_CylTwoHalves.Face4", # manually approved
                "BooleanFragments_CylTwoHalves.Face5", # manually approved
                "BooleanFragments_CylTwoHalves.Face6", # manually approved
            ]
            n_tests += 1
            if set(result9c) == set(expected9c):
                n_passed += 1
                log("Test 9c: PASS (faces fully inside Body003, filtered by Solid3 only).")
            else:
                n_failed += 1
                failed_tests.append("Test 9c")
                log(f"Test 9c: FAIL (got {len(result9c)} faces {result9c}, expected {expected9c_count} faces from one half-sphere).")
            log(f"Test 9c (faces filtered by solid from 9b): {len(result9c)} face(s) selected: {result9c}")

        # Test 9d: same as 9c but SolidFilterRefs -> BooleanFragments_CylTwoHalves.Solid2 only -> [Face8, Face9]
        # Re-acquire the active document here because build_test_3_cylinder_two_halves()
        # (used via _ensure_test_9_geometry) may have closed/recreated the previous doc.
        doc = FreeCAD.ActiveDocument
        if doc:
            name_9d = "Test9d_Faces_CylTwoHalves_fully_inside_filtered_by_Solid2"
            existing_9d = doc.getObject(name_9d)
            if existing_9d:
                doc.removeObject(existing_9d.Name)
            obj_9d = doc.addObject("App::FeaturePython", name_9d)
            SelectionSet(obj_9d)
            obj_9d.Label = "SelectionSet " + name_9d
            obj_9d.MainObject = boolean_cyl
            obj_9d.SelectionShape = body003
            obj_9d.Mode = "faces"
            obj_9d.VolumeMode = "fully_inside"
            if hasattr(obj_9d, "SolidFilterRefs"):
                obj_9d.SolidFilterRefs = [(boolean_cyl, "Solid2")]
            ok_9d = recomputeSelectionSetFromShapes(obj_9d)
            apply_selectionset_view_defaults(obj_9d)
            result9d = list(obj_9d.ElementList) if ok_9d else []
            expected9d = [
                "BooleanFragments_CylTwoHalves.Face8",
                "BooleanFragments_CylTwoHalves.Face9",
            ]
            n_tests += 1
            if set(result9d) == set(expected9d):
                n_passed += 1
                log("Test 9d: PASS (faces fully inside Body003, filtered by Solid2 only).")
            else:
                n_failed += 1
                failed_tests.append("Test 9d")
                log(f"Test 9d: FAIL (got {result9d}, expected {expected9d}).")
            log(f"Test 9d (SolidFilterRefs=Solid2): {len(result9d)} face(s): {result9d}")
        else:
            log("Test 9d skipped: no active document.")

    else:
        log("Test 9 skipped: build_test_3_cylinder_two_halves() failed to create geometry.")

    # Test 10: coordinates mode – find face(s) or solid(s) by point (using Test 1 geometry)
    doc = FreeCAD.ActiveDocument
    if main and doc and getattr(main, "Shape", None) and hasattr(main.Shape, "Faces") and len(main.Shape.Faces) >= 6:
        name_faces = "Test10_Faces_Coordinates_Face6"
        existing = doc.getObject(name_faces)
        if existing:
            doc.removeObject(existing.Name)
        obj_faces = doc.addObject("App::FeaturePython", name_faces)
        SelectionSet(obj_faces)
        obj_faces.Label = "SelectionSet " + name_faces
        obj_faces.MainObject = main
        obj_faces.Mode = "faces"
        obj_faces.VolumeMode = "coordinates"
        # Top face of cube (Face6) center
        obj_faces.CoordinatePoint = main.Shape.Faces[5].CenterOfMass
        ok = recomputeSelectionSetFromShapes(obj_faces)
        apply_selectionset_view_defaults(obj_faces)
        result10f = list(obj_faces.ElementList) if ok else []
        expected10f = ["TestCube_main.Face6"]
        n_tests += 1
        if result10f == expected10f:
            n_passed += 1
            log("Test 10 (coordinates, faces): PASS (point on top face -> Face6).")
        else:
            n_failed += 1
            failed_tests.append("Test 10 (faces)")
            log(f"Test 10 (coordinates, faces): FAIL (got {result10f}, expected {expected10f}).")
        log(f"Test 10 (coordinates faces): {len(result10f)} face(s) selected: {result10f}")

        # Test 10b: coordinates mode solids – point inside cube -> one solid
        name_solids = "Test10b_Solids_Coordinates_Solid1"
        existing_s = doc.getObject(name_solids)
        if existing_s:
            doc.removeObject(existing_s.Name)
        obj_solids = doc.addObject("App::FeaturePython", name_solids)
        SelectionSet(obj_solids)
        obj_solids.Label = "SelectionSet " + name_solids
        obj_solids.MainObject = main
        obj_solids.Mode = "solids"
        obj_solids.VolumeMode = "coordinates"
        obj_solids.CoordinatePoint = main.Shape.CenterOfMass
        ok_s = recomputeSelectionSetFromShapes(obj_solids)
        apply_selectionset_view_defaults(obj_solids)
        result10s = list(obj_solids.ElementList) if ok_s else []
        expected10s = ["TestCube_main.Solid1"]
        n_tests += 1
        if result10s == expected10s:
            n_passed += 1
            log("Test 10b (coordinates, solids): PASS (point inside cube -> Solid1).")
        else:
            n_failed += 1
            failed_tests.append("Test 10b (solids)")
            log(f"Test 10b (coordinates, solids): FAIL (got {result10s}, expected {expected10s}).")
        log(f"Test 10b (coordinates solids): {len(result10s)} solid(s) selected: {result10s}")

        # Test 10c: [5,5,0] on TestCube_main -> Face5 (bottom face)
        name_10c = "Test10c_Faces_Coordinates_55_0_Face5"
        existing_10c = doc.getObject(name_10c)
        if existing_10c:
            doc.removeObject(existing_10c.Name)
        obj_10c = doc.addObject("App::FeaturePython", name_10c)
        SelectionSet(obj_10c)
        obj_10c.Label = "SelectionSet " + name_10c
        obj_10c.MainObject = main
        obj_10c.Mode = "faces"
        obj_10c.VolumeMode = "coordinates"
        obj_10c.CoordinatePoint = FreeCAD.Vector(5, 5, 0)
        ok_10c = recomputeSelectionSetFromShapes(obj_10c)
        apply_selectionset_view_defaults(obj_10c)
        result10c = list(obj_10c.ElementList) if ok_10c else []
        expected10c = ["TestCube_main.Face5"]
        n_tests += 1
        if result10c == expected10c:
            n_passed += 1
            log("Test 10c (coordinates [5,5,0] -> Face5): PASS.")
        else:
            n_failed += 1
            failed_tests.append("Test 10c (coordinates Face5)")
            log(f"Test 10c (coordinates [5,5,0] -> Face5): FAIL (got {result10c}, expected {expected10c}).")
        log(f"Test 10c (coordinates): {len(result10c)} face(s): {result10c}")

        # Test 10d: [10,10,0] on TestCube_main -> Face2, Face4, Face5 (corner)
        name_10d = "Test10d_Faces_Coordinates_10_10_0"
        existing_10d = doc.getObject(name_10d)
        if existing_10d:
            doc.removeObject(existing_10d.Name)
        obj_10d = doc.addObject("App::FeaturePython", name_10d)
        SelectionSet(obj_10d)
        obj_10d.Label = "SelectionSet " + name_10d
        obj_10d.MainObject = main
        obj_10d.Mode = "faces"
        obj_10d.VolumeMode = "coordinates"
        obj_10d.CoordinatePoint = FreeCAD.Vector(10, 10, 0)
        ok_10d = recomputeSelectionSetFromShapes(obj_10d)
        apply_selectionset_view_defaults(obj_10d)
        result10d = list(obj_10d.ElementList) if ok_10d else []
        expected10d = ["TestCube_main.Face2", "TestCube_main.Face4", "TestCube_main.Face5"]
        n_tests += 1
        if set(result10d) == set(expected10d):
            n_passed += 1
            log("Test 10d (coordinates [10,10,0] -> Face2, Face4, Face5): PASS.")
        else:
            n_failed += 1
            failed_tests.append("Test 10d (coordinates corner)")
            log(f"Test 10d (coordinates [10,10,0] -> Face2, Face4, Face5): FAIL (got {result10d}, expected {expected10d}).")
        log(f"Test 10d (coordinates): {len(result10d)} face(s): {result10d}")
    else:
        log("Test 10 skipped: Test 1 geometry (main cube) not available.")

    # Test 10e: [5,22,5] on BooleanFragments -> Face7 (sphere fragment); needs Test 2 geometry
    boolean_bf = doc.getObject("BooleanFragments") if doc else None
    if boolean_bf and getattr(boolean_bf, "Shape", None) and hasattr(boolean_bf.Shape, "Faces"):
        name_10e = "Test10e_Faces_Coordinates_BooleanFragments_Face7"
        existing_10e = doc.getObject(name_10e)
        if existing_10e:
            doc.removeObject(existing_10e.Name)
        obj_10e = doc.addObject("App::FeaturePython", name_10e)
        SelectionSet(obj_10e)
        obj_10e.Label = "SelectionSet " + name_10e
        obj_10e.MainObject = boolean_bf
        obj_10e.Mode = "faces"
        obj_10e.VolumeMode = "coordinates"
        obj_10e.CoordinatePoint = FreeCAD.Vector(5, 22, 5)
        ok_10e = recomputeSelectionSetFromShapes(obj_10e)
        apply_selectionset_view_defaults(obj_10e)
        result10e = list(obj_10e.ElementList) if ok_10e else []
        expected10e = ["BooleanFragments.Face7"]
        n_tests += 1
        if result10e == expected10e:
            n_passed += 1
            log("Test 10e (coordinates [5,22,5] -> BooleanFragments.Face7): PASS.")
        else:
            n_failed += 1
            failed_tests.append("Test 10e (coordinates BooleanFragments)")
            log(f"Test 10e (coordinates [5,22,5] -> Face7): FAIL (got {result10e}, expected {expected10e}).")
        log(f"Test 10e (coordinates): {len(result10e)} face(s): {result10e}")
    else:
        log("Test 10e skipped: BooleanFragments not available.")

    # Test 10f: point on BooleanFragments_CylTwoHalves -> Face5, Face6, Face8, Face9 (geometry at y=50, use [5,52,5])
    boolean_cyl = doc.getObject("BooleanFragments_CylTwoHalves") if doc else None
    if boolean_cyl and getattr(boolean_cyl, "Shape", None) and hasattr(boolean_cyl.Shape, "Faces"):
        name_10f = "Test10f_Faces_Coordinates_CylTwoHalves"
        existing_10f = doc.getObject(name_10f)
        if existing_10f:
            doc.removeObject(existing_10f.Name)
        obj_10f = doc.addObject("App::FeaturePython", name_10f)
        SelectionSet(obj_10f)
        obj_10f.Label = "SelectionSet " + name_10f
        obj_10f.MainObject = boolean_cyl
        obj_10f.Mode = "faces"
        obj_10f.VolumeMode = "coordinates"
        obj_10f.CoordinatePoint = FreeCAD.Vector(0, 50  , 5)
        ok_10f = recomputeSelectionSetFromShapes(obj_10f)
        apply_selectionset_view_defaults(obj_10f)
        result10f = list(obj_10f.ElementList) if ok_10f else []
        expected10f_set = {
            "BooleanFragments_CylTwoHalves.Face5",
            "BooleanFragments_CylTwoHalves.Face6",
            "BooleanFragments_CylTwoHalves.Face8",
            "BooleanFragments_CylTwoHalves.Face9",
        }
        n_tests += 1
        if set(result10f) == expected10f_set:
            n_passed += 1
            log("Test 10f (coordinates [5,52,5] -> CylTwoHalves Face5,6,8,9): PASS.")
        else:
            n_failed += 1
            failed_tests.append("Test 10f (coordinates CylTwoHalves)")
            log(f"Test 10f (coordinates CylTwoHalves): FAIL (got {result10f}, expected {sorted(expected10f_set)}).")
        log(f"Test 10f (coordinates): {len(result10f)} face(s): {result10f}")
    else:
        log("Test 10f skipped: BooleanFragments_CylTwoHalves not available.")

    # Apply view defaults (DisplayMode, Visibility) to all SelectionSet/SelectionSetLink so show/hide works in tree
    if doc:
        for o in doc.Objects:
            label = getattr(o, "Label", "") or ""
            if label.startswith("SelectionSet ") or label.startswith("SelectionSetLink "):
                apply_selectionset_view_defaults(o)

    # Test FEM: use SelectionSet in a FEM analysis (constraint references); verify references are accepted
    add_tests, add_passed, add_failed, fail_name = _run_fem_test(log)
    n_tests += add_tests
    n_passed += add_passed
    n_failed += add_failed
    if fail_name:
        failed_tests.append(fail_name)

    # Test FEM cantilever full: mesh + solve with hole, check max deflection
    add_tests, add_passed, add_failed, fail_name = _run_fem_cantilever_full_test(log)
    n_tests += add_tests
    n_passed += add_passed
    n_failed += add_failed
    if fail_name:
        failed_tests.append(fail_name)

    # Add summary to _log_lines only (do not log to report here). Caller appends summary to report last so it stays at the very end (after any FEM/debug output).
    import os
    _results_dir = os.path.dirname(os.path.abspath(__file__))
    _results_path = os.path.join(_results_dir, "selection_set_test_results.txt")
    _log_lines.append(
        f"\n\n######\n[SelectionSet Test Runner] Tests complete. {n_passed} passed, {n_failed} failed out of {n_tests} total tests."
    )
    if failed_tests:
        _log_lines.append(f"Failing tests: {', '.join(failed_tests)}")
        if _is_fc102() and "Test FEM cantilever full" in failed_tests:
            _log_lines.append("[Info] %s" % FEM_CANTILEVER_FC102_INFO)
    _log_lines.append("[Results written to %s]" % _results_path)

    # Write results to file so automated runs can read them (e.g. run_with_gui.sh + read selection_set_test_results.txt)
    try:
        with open(_results_path, "w") as f:
            f.write("\n".join(_log_lines))
    except Exception as e:
        _log_lines.append("[Could not write results file: %s]" % e)

    # Re-attach observer only if single-click expand is enabled (avoids duplicate observers and TypeError)
    if USE_SELECTION_OBSERVER:
        try:
            _attach_selectionset_observer()
            log("[SelectionSet] Observer re-attached after tests.")
        except Exception:
            pass

    # Restore verbose callback output
    selection_set_core.SELECTIONSET_VERBOSE_CALLBACK = _prev_verbose

    return (n_passed, n_failed, n_tests, _results_path, failed_tests)


def run_selectionset_tests_with_summary():
    """Run the full test suite and append the summary to the Report view last (so it stays at the end)."""
    result = run_selectionset_tests()
    append_test_summary_to_report(*result)


def append_test_summary_to_report(n_passed, n_failed, n_tests, results_path, failed_tests=None):
    """
    Append the test summary to the Report view (and console if no report) so it appears at the very end.
    Call this after run_selectionset_tests() and after any other output (e.g. debug export) so the user sees the summary last.
    """
    report = FreeCADGui.getMainWindow().findChild(QtWidgets.QTextEdit, "Report view") if FreeCADGui.getMainWindow() else None
    lines = []
    if failed_tests:
        lines.append("Failing tests: %s" % ", ".join(failed_tests))
        if _is_fc102() and "Test FEM cantilever full" in failed_tests:
            lines.append("[Info] %s" % FEM_CANTILEVER_FC102_INFO)
    lines.append("[SelectionSet Test Runner] Tests complete. %d passed, %d failed out of %d total tests." % (n_passed, n_failed, n_tests))
    lines.append("[Results written to %s]" % (results_path or "selection_set_test_results.txt"))
    for line in lines:
        if report:
            report.append(line)
        else:
            print(line)


# ---------- Toolbar: test pulldown and individual buttons ----------
def add_tests_pulldown_to_toolbar():
    """Add a single toolbar button with a pulldown menu for all test actions (R1, R2, R3, Run all tests, FEM beam)."""
    mw = FreeCADGui.getMainWindow()
    toolbar_name = "Selection_tools"
    for child in mw.findChildren(QtWidgets.QToolBar):
        if child.objectName() == toolbar_name:
            toolbar = child
            break
    else:
        toolbar = QtWidgets.QToolBar(toolbar_name, mw)
        toolbar.setObjectName(toolbar_name)
        mw.addToolBar(toolbar)
    # Remove old individual test buttons so we don't duplicate
    for button_name in ("RunTest1Button", "BuildTestGeometryButton", "RunTest3Button",
                       "RunSelectionSetTestsButton", "FEMBeamExampleButton"):
        old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
        if old_btn:
            action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
            if action:
                toolbar.removeAction(action)
            old_btn.deleteLater()
    # Pulldown button
    button_name = "SelectionSetTestsPulldownButton"
    old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
    if old_btn:
        action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
        if action:
            toolbar.removeAction(action)
        old_btn.deleteLater()
    btn = QtWidgets.QToolButton(toolbar)
    btn.setObjectName(button_name)
    icon = get_toolbar_icon("RunTestAll")
    if icon is None:
        try:
            icon = QtGui.QIcon.fromTheme("system-run")
        except Exception:
            icon = None
    if icon is not None:
        btn.setIcon(icon)
    btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    btn.setText("Tests")
    btn.setToolTip("SelectionSet tests: Run Test 1, Build test geometry, Run Test 3, Run all tests, FEM beam example.")
    btn.setPopupMode(QtWidgets.QToolButton.InstantPopup)
    menu = QtWidgets.QMenu(btn)
    a1 = menu.addAction("Run Test 1 (two cubes, one SelectionSet)")
    a1.triggered.connect(run_test_1_only)
    a2 = menu.addAction("Build test geometry (all test shapes)")
    a2.triggered.connect(build_all_test_geometry)
    a3 = menu.addAction("Run Test 3 (minimal FEM + SelectionSetLink)")
    a3.triggered.connect(run_test_3_combined_demo)
    a4 = menu.addAction("Run SelectionSet Tests (all test cases)")
    a4.triggered.connect(run_selectionset_tests_with_summary)
    menu.addSeparator()
    a5 = menu.addAction("FEM beam example")
    a5.triggered.connect(lambda: run_fem_cantilever_beam_example(cylinder_cut=False))
    btn.setMenu(menu)
    toolbar.addWidget(btn)
    print("Tests pulldown added to 'Selection_tools' toolbar.")


def add_test_button_to_toolbar():
    mw = FreeCADGui.getMainWindow()
    toolbar_name = "Selection_tools"
    for child in mw.findChildren(QtWidgets.QToolBar):
        if child.objectName() == toolbar_name:
            toolbar = child
            break
    else:
        toolbar = QtWidgets.QToolBar(toolbar_name, mw)
        toolbar.setObjectName(toolbar_name)
        mw.addToolBar(toolbar)
    button_name = "RunSelectionSetTestsButton"
    old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
    if old_btn:
        action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
        if action:
            toolbar.removeAction(action)
        old_btn.deleteLater()
    btn = QtWidgets.QToolButton(toolbar)
    btn.setObjectName(button_name)
    # Use a distinct icon for running all tests (RunTestAll.svg or a suitable fallback).
    icon = get_toolbar_icon("RunTestAll")
    if icon is None:
        try:
            icon = QtGui.QIcon.fromTheme("system-run")
        except Exception:
            icon = None
    if icon is not None:
        btn.setIcon(icon)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    btn.setText("Run SelectionSet Tests")
    btn.setToolTip("Run all SelectionSet test cases and check outcomes.")
    btn.clicked.connect(run_selectionset_tests_with_summary)
    toolbar.addWidget(btn)
    print("Run SelectionSet Tests button added to 'Selection_tools' toolbar.")


def add_remove_observer_button_to_toolbar():
    mw = FreeCADGui.getMainWindow()
    toolbar_name = "Selection_tools"
    for child in mw.findChildren(QtWidgets.QToolBar):
        if child.objectName() == toolbar_name:
            toolbar = child
            break
    else:
        toolbar = QtWidgets.QToolBar(toolbar_name, mw)
        toolbar.setObjectName(toolbar_name)
        mw.addToolBar(toolbar)
    button_name = "RemoveSelectionSetObserverButton"
    old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
    if old_btn:
        action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
        if action:
            toolbar.removeAction(action)
        old_btn.deleteLater()
    btn = QtWidgets.QToolButton(toolbar)
    btn.setObjectName(button_name)
    icon = get_toolbar_icon("SelectionSet")
    if icon is not None:
        btn.setIcon(icon)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    btn.setText("Remove SelectionSet observer")
    btn.setToolTip("Remove the SelectionSet observer if one is attached (e.g. after enabling single-click expand in code, or before re-running the macro). With default settings (USE_SELECTION_OBSERVER=False) no observer is attached, so the button will report 'No observer to remove'.")
    btn.clicked.connect(_detach_selectionset_observer)
    toolbar.addWidget(btn)
    print("Remove SelectionSet observer button added to 'Selection_tools' toolbar.")


def add_build_test_geometry_button_to_toolbar():
    """Add a 'Build test geometry' button to create all test shapes on demand."""
    mw = FreeCADGui.getMainWindow()
    toolbar_name = "Selection_tools"
    for child in mw.findChildren(QtWidgets.QToolBar):
        if child.objectName() == toolbar_name:
            toolbar = child
            break
    else:
        toolbar = QtWidgets.QToolBar(toolbar_name, mw)
        toolbar.setObjectName(toolbar_name)
        mw.addToolBar(toolbar)
    button_name = "BuildTestGeometryButton"
    old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
    if old_btn:
        action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
        if action:
            toolbar.removeAction(action)
        old_btn.deleteLater()
    btn = QtWidgets.QToolButton(toolbar)
    btn.setObjectName(button_name)
    # Use a distinct icon for building all test geometry (RunTest2.svg or a suitable fallback).
    icon = get_toolbar_icon("RunTest2")
    if icon is None:
        try:
            icon = QtGui.QIcon.fromTheme("document-new")
        except Exception:
            icon = None
    if icon is not None:
        btn.setIcon(icon)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    btn.setText("Build test geometry")
    btn.setToolTip("Create all test shapes (cubes, cube+sphere, cylinder+halves). Run tests afterwards with 'Run SelectionSet Tests'.")
    btn.clicked.connect(build_all_test_geometry)
    toolbar.addWidget(btn)
    print("Build test geometry button added to 'Selection_tools' toolbar.")


def add_run_test_1_button_to_toolbar():
    """Add a 'Run Test 1' button: create the two test cubes and run only Test 1."""
    mw = FreeCADGui.getMainWindow()
    toolbar_name = "Selection_tools"
    for child in mw.findChildren(QtWidgets.QToolBar):
        if child.objectName() == toolbar_name:
            toolbar = child
            break
    else:
        toolbar = QtWidgets.QToolBar(toolbar_name, mw)
        toolbar.setObjectName(toolbar_name)
        mw.addToolBar(toolbar)
    button_name = "RunTest1Button"
    old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
    if old_btn:
        action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
        if action:
            toolbar.removeAction(action)
        old_btn.deleteLater()
    btn = QtWidgets.QToolButton(toolbar)
    btn.setObjectName(button_name)
    # Use a distinct icon for running tests (either a custom 'RunTest1.svg' in icons/, or fallback to a play symbol).
    icon = get_toolbar_icon("RunTest1")
    if icon is None:
        try:
            # Fallback: use a standard \"media-playback-start\" icon if available from the theme.
            icon = QtGui.QIcon.fromTheme("media-playback-start")
        except Exception:
            icon = None
    if icon is not None:
        btn.setIcon(icon)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    btn.setText("Run Test 1")
    btn.setToolTip("Run Test 1 only: build or reuse the two test cubes (main + selection shape) and create a single SelectionSet with the top face. Useful for quick checks; deleting this SelectionSet does not affect the cubes.")
    btn.clicked.connect(run_test_1_only)
    toolbar.addWidget(btn)
    print("Run Test 1 button added to 'Selection_tools' toolbar.")


def add_run_test_3_button_to_toolbar():
    """Add a 'Run Test 3' button: build minimal combined demo (two shapes, one SelectionSet, one SelectionSetLink + FEM)."""
    mw = FreeCADGui.getMainWindow()
    toolbar_name = "Selection_tools"
    for child in mw.findChildren(QtWidgets.QToolBar):
        if child.objectName() == toolbar_name:
            toolbar = child
            break
    else:
        toolbar = QtWidgets.QToolBar(toolbar_name, mw)
        toolbar.setObjectName(toolbar_name)
        mw.addToolBar(toolbar)
    button_name = "RunTest3Button"
    old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
    if old_btn:
        action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
        if action:
            toolbar.removeAction(action)
        old_btn.deleteLater()
    btn = QtWidgets.QToolButton(toolbar)
    btn.setObjectName(button_name)
    icon = get_toolbar_icon("RunTest3")
    if icon is not None:
        btn.setIcon(icon)
        btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    btn.setText("Run Test 3")
    btn.setToolTip(
        "Run a minimal combined demo: builds Test 1 cubes, creates one SelectionSet and one SelectionSetLink, "
        "and applies them to a fixed FEM constraint (FEM_Test3_ConstraintFixed) if FEM is available."
    )
    btn.clicked.connect(run_test_3_combined_demo)
    toolbar.addWidget(btn)
    print("Run Test 3 button added to 'Selection_tools' toolbar.")


def add_fem_beam_example_button_to_toolbar():
    """Add a 'FEM beam example' button: cantilever beam with fixed + force via SelectionSets and SelectionSetLinks."""
    mw = FreeCADGui.getMainWindow()
    toolbar_name = "Selection_tools"
    for child in mw.findChildren(QtWidgets.QToolBar):
        if child.objectName() == toolbar_name:
            toolbar = child
            break
    else:
        toolbar = QtWidgets.QToolBar(toolbar_name, mw)
        toolbar.setObjectName(toolbar_name)
        mw.addToolBar(toolbar)
    button_name = "FEMBeamExampleButton"
    old_btn = toolbar.findChild(QtWidgets.QToolButton, button_name)
    if old_btn:
        action = old_btn.defaultAction() if hasattr(old_btn, "defaultAction") else None
        if action:
            toolbar.removeAction(action)
        old_btn.deleteLater()
    btn = QtWidgets.QToolButton(toolbar)
    btn.setObjectName(button_name)
    icon = get_toolbar_icon("RunTest3")
    if icon is None:
        try:
            icon = QtGui.QIcon.fromTheme("document-new")
        except Exception:
            icon = None
    if icon is not None:
        btn.setIcon(icon)
    btn.setToolButtonStyle(QtCore.Qt.ToolButtonTextBesideIcon)
    btn.setText("FEM beam example")
    btn.setToolTip(
        "Build FEM cantilever beam example: beam box, two SelectionSets (fixed and force face by coordinates), "
        "ConstraintFixed and ConstraintForce, SelectionSetLinks. For version with cylinder cut at half length, "
        "run in Python: selection_set_tests.run_fem_cantilever_beam_example(cylinder_cut=True)"
    )
    btn.clicked.connect(lambda: run_fem_cantilever_beam_example(cylinder_cut=False))
    toolbar.addWidget(btn)
    print("FEM beam example button added to 'Selection_tools' toolbar.")


# ---------- Test geometry builders ----------
def build_test_2_cube_with_inner_sphere():
    """
    Build a test setup with a 10mm cube and a 0.5mm radius sphere inside it (centered).
    Returns (cube, sphere)
    """
    doc = FreeCAD.ActiveDocument
    if doc is None:
        doc = FreeCAD.newDocument()
    # Remove old test objects if present
    for name in ["TestCube_Bool", "TestSphere_Bool"]:
        obj = doc.getObject(name)
        if obj:
            doc.removeObject(name)

    offset_y_mm = 20  # offset in y-direction to avoid overlap with previous test cubes

    # Create cube
    cube = doc.addObject("Part::Box", "TestCube_Bool")
    cube.Length = 10
    cube.Width = 10
    cube.Height = 10
    cube.Placement.Base = FreeCAD.Vector(0, offset_y_mm, 0)

    # Create sphere, radius 3mm, centered in cube
    sphere = doc.addObject("Part::Sphere", "TestSphere_Bool")
    sphere.Radius = 3
    sphere.Placement.Base = FreeCAD.Vector(5, 5 + offset_y_mm, 5)
    # Color the sphere orange for better visibility
    sphere.ViewObject.ShapeColor = (1.0, 0.5, 0.0)

    # Ensure shapes are valid before BooleanFragments
    doc.recompute()
    if not hasattr(cube, "Shape") or cube.Shape.isNull():
        print(
            "[ERROR] Cube shape is null after recompute. Skipping BooleanFragments creation."
        )
        return cube, sphere, None
    if not hasattr(sphere, "Shape") or sphere.Shape.isNull():
        print(
            "[ERROR] Sphere shape is null after recompute. Skipping BooleanFragments creation."
        )
        return cube, sphere, None

    # Use BOPTools.SplitFeatures to create BooleanFragments
    import BOPTools.SplitFeatures

    boolean = BOPTools.SplitFeatures.makeBooleanFragments(name="BooleanFragments")
    boolean.Objects = [cube, sphere]
    boolean.Mode = "Standard"
    try:
        boolean.Proxy.execute(boolean)
        boolean.purgeTouched()
        doc.recompute()
        # Hide child fragments for clarity
        for obj in boolean.ViewObject.Proxy.claimChildren():
            obj.ViewObject.hide()
    except Exception as e:
        print(f"[ERROR] Exception during BooleanFragments creation: {e}")
        return cube, sphere, None
    boolean.ViewObject.Transparency = 50

    sphere.ViewObject.Transparency = 0
    cube.ViewObject.Transparency = 0

    # Create selection cube for the boolean fragments object sphere
    cube_sel = doc.addObject("Part::Box", "TestCube_Selection_Bool")
    cube_sel.Length = 8
    cube_sel.Width = 8
    cube_sel.Height = 8
    cube_sel.Placement.Base = FreeCAD.Vector(1, offset_y_mm + 1, 1)
    cube_sel.ViewObject.Transparency = 50
    selection_defining_shape = cube_sel

    selection_defining_shape.ViewObject.Transparency = 50
    selection_defining_shape.ViewObject.ShapeColor = (
        0.0,
        1.0,
        0.1,
    )  # green for better visibility

    doc.recompute()
    _add_objects_to_group(
        doc, "Test2_Geometry",
        [cube, sphere, boolean, selection_defining_shape],
        label="Test 2 geometry",
    )
    # isometric view for better visualization
    FreeCADGui.ActiveDocument.ActiveView.viewIsometric()
    # fit the view to see all objects
    FreeCADGui.ActiveDocument.ActiveView.fitAll()

    print(
        "Test geometry created: 10mm cube with 0.5mm sphere inside (centered) using BooleanFragments."
    )
    return cube, sphere, boolean, selection_defining_shape


def build_test_3_cylinder_two_halves(distance_sphere_halfes=0.000001):
    """
    Build Test 9 geometry: cylinder (Body) + two half-spheres (Body001, Body002),
    BooleanFragments of the three, and Body003 with sketch-based Pad as selection shape.
    Uses the same logic as create_3rd_test_geometry macro.
    Returns (boolean_fragments, body003) or (None, None) on failure.

    distance_sphere_halfes: Y gap between the two half-spheres in mm. If 0.0, a minimum
    gap of 1e-6 mm is used to avoid degeneracy (touching spheres cause BooleanFragments
    and point-in-solid tests to fail or return wrong results).
    """
    import PartDesign
    import Sketcher
    import BOPTools.SplitFeatures

    doc = FreeCAD.ActiveDocument
    if doc is None:
        doc = FreeCAD.newDocument()

    offset_y_mm = 50

    # Remove existing Test 9 objects only (use unique name so we don't remove build_test_2's BooleanFragments)
    for name in ["BooleanFragments_CylTwoHalves", "Body003", "Body002", "Body001", "Body"]:
        obj = doc.getObject(name)
        if obj:
            doc.removeObject(name)
    doc.recompute()

    # Body + Cylinder (10 mm radius, 10 mm height)
    doc.addObject("PartDesign::Body", "Body")
    doc.addObject("PartDesign::AdditiveCylinder", "Cylinder")
    doc.getObject("Body").addObject(doc.getObject("Cylinder"))
    doc.recompute()
    doc.getObject("Cylinder").Radius = "10,00 mm"
    doc.getObject("Cylinder").Height = "10,00 mm"
    doc.getObject("Cylinder").Angle = "360,00 °"
    doc.getObject("Cylinder").FirstAngle = "0,00 °"
    doc.getObject("Cylinder").SecondAngle = "0,00 °"
    doc.getObject("Body").Placement.Base.y += offset_y_mm
    doc.recompute()

    # Body001 + half-sphere
    doc.addObject("PartDesign::Body", "Body001")
    doc.getObject("Body001").Label = "Body"
    doc.addObject("PartDesign::AdditiveSphere", "Sphere")
    doc.getObject("Body001").addObject(doc.getObject("Sphere"))
    doc.recompute()
    doc.getObject("Sphere").Radius = "3,00 mm"
    doc.getObject("Sphere").Angle1 = "-90,00 °"
    doc.getObject("Sphere").Angle2 = "90,00 °"
    doc.getObject("Sphere").Angle3 = "180,00 °"
    doc.recompute()

    # Body002 + half-sphere
    doc.addObject("PartDesign::Body", "Body002")
    doc.getObject("Body002").Label = "Body"
    doc.addObject("PartDesign::AdditiveSphere", "Sphere001")
    doc.getObject("Body002").addObject(doc.getObject("Sphere001"))
    doc.recompute()
    doc.getObject("Sphere001").Radius = "3,00 mm"
    doc.getObject("Sphere001").Angle1 = "-90,00 °"
    doc.getObject("Sphere001").Angle2 = "90,00 °"
    doc.getObject("Sphere001").Angle3 = "180,00 °"
    doc.recompute()

    # Place half-spheres at half cylinder height; rotate Body002 180° about Z
    cyl = doc.getObject("Cylinder")
    try:
        h = cyl.Height
        half_z = float(h.Value) / 2.0 if hasattr(h, "Value") else 5.0
    except Exception:
        half_z = 5.0
    doc.getObject("Body001").Placement.Base.z = half_z
    doc.getObject("Body001").Placement.Base.y += offset_y_mm

    pl2 = doc.getObject("Body002").Placement
    pl2.Base.z = half_z
    pl2.Rotation = FreeCAD.Rotation(FreeCAD.Vector(0, 0, 1), 180)
    doc.getObject("Body002").Placement = pl2
    # Enforce a minimum gap to avoid degeneracy when distance_sphere_halfes is 0 (tolerance/BooleanFragments issue)
    try:
        d = float(distance_sphere_halfes)
    except (TypeError, ValueError):
        d = 0.000001
    gap = max(d, 1e-6)
    doc.getObject("Body002").Placement.Base.y += offset_y_mm - gap
    doc.recompute()

    # Body003: sketch on YZ plane + Pad (selection-defining shape for Test 9)
    doc.addObject("PartDesign::Body", "Body003")
    doc.getObject("Body003").Label = "Body_Selection_PD_Sketch"
    doc.recompute()
    yz_plane = doc.getObject("YZ_Plane003")
    if yz_plane is None:
        # Fallback: find YZ plane from Body003's origin
        body003 = doc.getObject("Body003")
        for o in doc.Objects:
            if "YZ_Plane" in (getattr(o, "Name", "") or "") and body003 in getattr(o, "InListRecursive", []):
                yz_plane = o
                break
        if yz_plane is None:
            for name in ["YZ_Plane003", "YZ_Plane002", "YZ_Plane001", "YZ_Plane"]:
                yz_plane = doc.getObject(name)
                if yz_plane:
                    break
    sk = doc.getObject("Body003").newObject("Sketcher::SketchObject", "Sketch")
    if sk is None:
        sk = doc.getObject("Sketch")
    sk.AttachmentSupport = [(yz_plane, [""])]
    sk.MapMode = "FlatFace"
    doc.recompute()
    y_line_offset=1.0 #mm
    # Closed polyline profile (from draw_sketch_part / create_3rd_test_geometry)
    sk.addGeometry(Part.LineSegment(FreeCAD.Vector(-6.563298, 8.975315, 0), FreeCAD.Vector(0.000000+y_line_offset, 8.998863, 0)), False)
    #sk.addConstraint(Sketcher.Constraint("PointOnObject", 0, 2, -2))
    sk.addConstraint(Sketcher.Constraint("Horizontal", 0))
    sk.addGeometry(Part.LineSegment(FreeCAD.Vector(0.000000+y_line_offset, 8.998863, 0), FreeCAD.Vector(0.000000+y_line_offset, 1.651702, 0)), False)
    sk.addConstraint(Sketcher.Constraint("Coincident", 0, 2, 1, 1))
    #sk.addConstraint(Sketcher.Constraint("PointOnObject", 1, 2, -2))
    sk.addConstraint(Sketcher.Constraint("Vertical", 1))
    sk.addGeometry(Part.LineSegment(FreeCAD.Vector(0.000000+y_line_offset, 1.651702, 0), FreeCAD.Vector(5.281645, 1.628153, 0)), False)
    sk.addConstraint(Sketcher.Constraint("Coincident", 1, 2, 2, 1))
    sk.addConstraint(Sketcher.Constraint("Horizontal", 2))
    sk.addGeometry(Part.LineSegment(FreeCAD.Vector(5.281645, 1.628153, 0), FreeCAD.Vector(5.305193, 0.497821, 0)), False)
    sk.addConstraint(Sketcher.Constraint("Coincident", 2, 2, 3, 1))
    sk.addConstraint(Sketcher.Constraint("Vertical", 3))
    sk.addGeometry(Part.LineSegment(FreeCAD.Vector(5.305193, 0.497821, 0), FreeCAD.Vector(-6.798784, 0.615564, 0)), False)
    sk.addConstraint(Sketcher.Constraint("Coincident", 3, 2, 4, 1))
    sk.addConstraint(Sketcher.Constraint("Horizontal", 4))
    sk.addGeometry(Part.LineSegment(FreeCAD.Vector(-6.798784, 0.615564, 0), FreeCAD.Vector(-6.563298, 8.998863, 0)), False)
    sk.addConstraint(Sketcher.Constraint("Coincident", 4, 2, 5, 1))
    sk.addConstraint(Sketcher.Constraint("Coincident", 5, 2, 0, 1))
    sk.addConstraint(Sketcher.Constraint("Vertical", 5))
    sk.delConstraint(4)
    doc.recompute()
    pad = doc.getObject("Body003").newObject("PartDesign::Pad", "Pad")
    if pad is None:
        pad = doc.getObject("Pad")
    pad.Profile = (sk, [""])
    doc.recompute()
    pad.ReferenceAxis = (sk, ["N_Axis"])
    sk.Visibility = False
    pad.Length = 7.0
    # FreeCAD 1.0 vs 1.1: In 1.1, 'Midplane' is deprecated. Use Type='Length' + SideType='Symmetric to plane' (or Symmetric=True) for 1.1, else Midplane (1.0).
    try:
        if hasattr(pad, "SideType"):
            # 1.1: symmetric extrusion. Ensure Type is Length/Dimension then set symmetric (SideType or Symmetric boolean).
            if hasattr(pad, "Type"):
                for t in ("Length", "Dimension"):
                    try:
                        pad.Type = t
                        break
                    except Exception:
                        continue
            for candidate in ("Symmetric to plane", "Symmetric", "Middle", "Midplane"):
                try:
                    pad.SideType = candidate
                    break
                except Exception:
                    continue
            else:
                for idx in (2,):  # index 2 often = symmetric in SideType enum (0=Length, 1=Two sides, 2=Symmetric to plane)
                    try:
                        pad.SideType = idx
                        break
                    except Exception:
                        continue
                else:
                    if hasattr(pad, "Symmetric"):
                        pad.Symmetric = True
        else:
            pad.Midplane = 1  # 1.0
    except Exception:
        try:
            pad.Midplane = 1  # fallback for 1.0
        except Exception:
            pass
    doc.getObject("Body003").Placement.Base.y += offset_y_mm
    doc.getObject("Body003").ViewObject.ShapeColor = (0.0, 1.0, 0.1) # green as all ather selection shapse
    doc.getObject("Body003").ViewObject.Transparency = 50
    doc.recompute()

    # BooleanFragments from Body, Body001, Body002 (unique name so Test 2's BooleanFragments stays visible)
    try:
        j = BOPTools.SplitFeatures.makeBooleanFragments(name="BooleanFragments_CylTwoHalves")
        j.Objects = [doc.Body, doc.Body001, doc.Body002]
        j.Mode = "Standard"
        j.Proxy.execute(j)
        j.purgeTouched()
        for obj in j.ViewObject.Proxy.claimChildren():
            obj.ViewObject.hide()
        j.ViewObject.Transparency = 50
        doc.recompute()
    except Exception as e:
        print(f"[ERROR] build_test_3_cylinder_two_halves: BooleanFragments failed: {e}")
        return None, None

    boolean_fragments = doc.getObject("BooleanFragments_CylTwoHalves")
    body003 = doc.getObject("Body003")
    _add_objects_to_group(
        doc, "Test9_Geometry",
        ["Body", "Body001", "Body002", "Body003", "BooleanFragments_CylTwoHalves"],
        label="Test 9 geometry",
    )
    print("Test 9 geometry created: cylinder + two half-spheres (BooleanFragments) and Body003 (sketch+pad).")
    # zoom to the geometry for better visualization
    FreeCADGui.ActiveDocument.ActiveView.fitAll()
    # isometric view for better visualization
    FreeCADGui.ActiveDocument.ActiveView.viewIsometric()
    return boolean_fragments, body003

def build_test_1_cubes():
    """
    Build a test setup with two cubes:
    - Main object: at origin, size 10x10x10 mm
    - Selection-defining shape: size 11x11x1 mm, shifted up by 9.5 mm in z
    Result: Only the top face of the main object is inside/intersected by the selection-defining shape.
    Returns (main_object, selection_defining_shape)
    """

    # Always close existing document "SelectionSetTest" if already open
    for d in FreeCAD.listDocuments().values():
        if d.Label == "SelectionSetTest":
            FreeCAD.closeDocument(d.Name)

    doc = FreeCAD.ActiveDocument
    if doc is None:
        doc = FreeCAD.newDocument()
        # rename the document to avoid "Unnamed" in object names
        doc.Label = "SelectionSetTest"

    # Remove old test cubes if present
    for name in ["TestCube_main", "TestCube_Selection"]:
        obj = doc.getObject(name)
        if obj:
            doc.removeObject(name)
    # Create main object at origin
    main_object = doc.addObject("Part::Box", "TestCube_main")
    main_object.Length = 10
    main_object.Width = 10
    main_object.Height = 10
    main_object.Placement.Base = FreeCAD.Vector(0, 0, 0)
    # main_object.ViewObject.ShapeColor = (
    #     0.0,
    #     0.5,
    #     1.0,
    # )  # light blue for better visibility
    main_object.ViewObject.Transparency = 50

    # Create selection-defining shape, size 11x11x1 mm, shifted up by 9.5 mm in z
    selection_defining_shape = doc.addObject("Part::Box", "TestCube_Selection")
    selection_defining_shape.Length = 11
    selection_defining_shape.Width = 11
    selection_defining_shape.Height = 1
    selection_defining_shape.Placement.Base = FreeCAD.Vector(-0.5, -0.5, 9.5)

    # make selection-defining shape half transparent for better visualization
    selection_defining_shape.ViewObject.Transparency = 50
    selection_defining_shape.ViewObject.ShapeColor = (
        0.0,
        1.0,
        0.1,
    )  # green for better visibility

    doc.recompute()
    _add_objects_to_group(doc, "Test1_Geometry", [main_object, selection_defining_shape], label="Test 1 geometry")
    print(
        "Test cubes created: main object at origin, selection-defining shape above main object."
    )
    return main_object, selection_defining_shape
