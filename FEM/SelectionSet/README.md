# FreeCAD macro: SelectionSet

Volume-based selection of faces or solids in FreeCAD. **SelectionSet:** define a “selection shape” (e.g. a box, sphere, or PartDesign body), then store which faces or solids of another object lie inside or intersect it as a **SelectionSet** in the model tree. **SelectionSetLink:** combine one or more SelectionSets (and/or direct shapes), add and subtract as needed, then apply the result to a FEM constraint, material, or any object with a References property—so you drive FEM geometry from your selections in one place.

---

## For users: using the macro

### How to run

1. In FreeCAD: **Macro → Macros → Execute**, then choose **FreeCAD_SelectionSet.py** (or open that file and run it).
2. A **Selection_tools** toolbar appears. For normal use you only need:
   - **Create/Update SelectionSet** – If only SelectionSet(s) are selected: recompute from stored Main/Shape. Otherwise: create a new SelectionSet or update selected sets from the current selection (main + shape; choose Faces or Solids).
   - **Expand SelectionSet** – Replace the GUI selection with the elements stored in the selected SelectionSet(s).
   - **Create SelectionSet link** – Add a link that combines SelectionSets/shapes and applies to one or more targets (e.g. FEM constraint, material). In the Data tab add **Target objects** (click …), set Add selection sets; then right-click the link → SelectionSet link → **Renew elements**.
3. In the **Data tab**, a SelectionSet has **MainObject**, **SelectionShape**, **Mode**, **VolumeMode** (fully_inside / intersection / coordinates), optional **CoordinatePoint** (for coordinates mode), and optional **SolidFilterRefs** (solids to filter by – click **...** to select solids like Main Object / Selection Shape). The Data tab only shows properties that are relevant for the current mode: **CoordinatePoint** appears only for VolumeMode = coordinates, and **SolidFilterRefs** only for face mode. Change the values, select the SelectionSet(s), and click **Create/Update SelectionSet** to refresh.
4. Double-click a SelectionSet in the tree to expand its elements in the selection and 3D view. Right-click shows a **SelectionSet** submenu with actions such as **Expand to elements (replace selection)** and **Delete SelectionSet only (keep shapes)**. **Shift+Click** selects only the main object; **Alt+Click** selects only the selection shape; **Ctrl+Click** expands to the elements (same as double-click). On a **SelectionSetLink**, **Ctrl+Click** expands to its combined elements.

Additional toolbar buttons (**Build test geometry**, **Run Test 1**, **Run Test 3**, **Run SelectionSet Tests**, **FEM beam example**) are for testing and development; see **For developers** below. **FEM beam example** builds a cantilever beam with fixed and force constraints driven by SelectionSets and SelectionSetLinks (coordinates mode for the end faces). For a version with a cylinder cut at half length, run in Python: `selection_set_tests.run_fem_cantilever_beam_example(cylinder_cut=True)`. Observer cleanup runs automatically when the macro is run, so re-running the macro does not create duplicate observers.

### Macro features (what you get)

- **Volume-based selection** – Main object's faces/solids filtered by a selection-defining shape: **fully inside** or **intersection** for faces; for solids, **fully_inside** = solid fully contained in the selection shape, **intersection** = any overlapping solid.
- **Selection by coordinates** – VolumeMode **coordinates**: find faces or solids that contain a point. Set **CoordinatePoint** (x,y,z) or **SelectionShape** (uses shape center); then select the set and click **Create/Update SelectionSet**. To hide the coordinate-point sphere in the 3D view: **Data** tab → **Display** → **ShowCoordinatePointMarker** (uncheck), or **View** tab → **Display** → **Show coordinate point** (uncheck). Works in FreeCAD 1.0 and 1.1.
- **Filter faces by solid** – Restrict to faces of specific solids: set **SolidFilterRefs** in the Data tab (click **...** to select solids, same as Main Object / Selection Shape), or select main + shape + a SelectionSet whose elements are solids. **How to use SolidFilterRefs:** In the Data tab, click the **...** button on the **SolidFilterRefs** row (tooltip: “Change the linked object”). In the link selection dialog, pick the main object (e.g. `BooleanFragments_CylTwoHalves`). In the **second row** of the dialog you can enter sub-elements: type e.g. `Solid2` or `Solid3` to restrict the selection to that solid. Add multiple entries (e.g. Solid1, Solid2) if needed. Confirm with OK, then select the set and click **Create/Update SelectionSet** to refresh the element list.
- **SelectionSet link** – Combine SelectionSets and/or direct shapes (add/subtract), apply to one or more **targets** (e.g. FEM constraint, material). In the Data tab set **Target objects** (click … to add); same refs are applied to each. Update targets only via **Renew elements** (right-click link → SelectionSet link). **Shape ref mode** = faces | solids; **UseOneByOneWorkaround** for one-face-at-a-time constraints.
- **FEM** – Use SelectionSets as references for FEM constraints/materials. For material geometry use **Face/Edge** (Solid not supported in FreeCAD's FEM material selector). Double-click a SelectionSet with a constraint selected to transfer refs.
- **Naming** – New SelectionSets get label **SelectionSet …**; new SelectionSetLinks get **SelectionSetLink …** with ref summary (e.g. "6 faces").

---

## For developers: tests and file layout

### File layout

| Role | Files |
|------|--------|
| **Macro (what users run)** | `FreeCAD_SelectionSet.py`, `selection_set_core.py`, `highlight_subelements.py` |
| **Tests and scripts** | `selection_set_tests.py`, `run_tests_headless.py`, `run_with_gui.sh`, `run_headless.sh`, `run_macro_and_tests_at_startup.py` / `.FCMacro` |
| **Docs and config** | `README.md`, `RELEASE.md`, `icons/`, `package.xml` (for Addon Manager). Not for addon: `rules.md`, `TODO.md` (dev only; in `.gitignore`). |

The macro does **not** depend on the test module for normal operation. The toolbar loads the test runner only to show optional test buttons.

### Python API (scripting without GUI selection)

All functionality is available from Python. Ensure the macro is loaded (e.g. run `FreeCAD_SelectionSet.py` once) and import `selection_set_core`.

| Purpose | Function |
|--------|----------|
| **Faces at a point** | `get_faces_by_point(obj, point, tol=1e-6, solid_filter_list=None)` → list of `"Obj.FaceN"` |
| **Solids at a point** | `get_solids_by_point(obj, point, tol=1e-6)` → list of `"Obj.SolidN"` |
| **Faces/solids by shape** | `get_faces_by_selection_shape(obj, shape, tol=1e-7, volume_mode="fully_inside", solid_filter_list=None)`; for solids use same helpers internally with `volume_mode` "fully_inside" or "intersection" |
| **Create SelectionSet** | `create_selectionset(doc, name, main_object, selection_shape=None, coordinate_point=None, mode="faces", volume_mode="fully_inside", solid_filter_refs=None, label=None)` → new SelectionSet (no GUI selection) |
| **Update from shapes** | `recomputeSelectionSetFromShapes(set_obj)` after changing MainObject / SelectionShape / CoordinatePoint |
| **Create link** | `createSelectionSetLink(name=None, _target=..., _add_sets=[], _add_shapes=[], _targets=[])`; use `_targets` for multiple targets |
| **Apply to FEM** | `get_fem_references_from_selectionset(set_obj)` → `[(obj, subname), ...]`; `apply_selectionset_to_fem_constraint(set_obj, constraint)`; `apply_references_to_fem_constraint_one_by_one(refs, constraint, doc)` for one-face-at-a-time constraints |
| **View defaults** | `apply_selectionset_view_defaults(obj)` after creating a set/link from script |
| **Create from GUI selection** | `createSelectionSetFromCurrent(selection_set_name=None, force_mode=None, volume_mode=None)` (requires current selection) |

Example (coordinates mode, point on a face):

```python
import FreeCAD
import selection_set_core
doc = FreeCAD.ActiveDocument
box = doc.getObject("Box")
pt = FreeCAD.Vector(5, 5, 10)  # e.g. top face center
ss = selection_set_core.create_selectionset(
    doc=doc, name="MySet", main_object=box,
    coordinate_point=pt, mode="faces", volume_mode="coordinates"
)
# ss.ElementList now holds the face(s) containing the point
```

### Running the test suite

- **Inside FreeCAD:** Run the macro, then click **Run SelectionSet Tests**. Results in Report view and **selection_set_test_results.txt**.
- **With startup script (e.g. CI):** Put a FreeCAD AppImage in this folder, run `./run_with_gui.sh` or `./run_with_gui.sh --exit`. Results in **selection_set_test_results.txt**.
- **Headless:** `./run_headless.sh` or `freecad -c FreeCAD_macro_SelectionSet/run_tests_headless.py`. Output in **test9c_headless_output.txt** (for debugging Test 9c).

Tests: 1–9, 9a, 9b, **9c** (faces filtered by Solid3), **9d** (faces filtered by Solid2), **10** (coordinates faces), **10b** (coordinates solids), **10c–10f** (coordinates at specific points), **Test FEM**, **Test FEM cantilever full** (with hole r=300, no hole, hole r=400; skipped if Fem not loaded; on 1.0.2 GMSH with hole may fail).

### Test / development rules

- **rules.md:** Do not change the values of `expected*` variables in the tests without asking.
- **Debug Test 9c:** Use `./run_headless.sh` or `freecad -c FreeCAD_macro_SelectionSet/run_tests_headless.py` and inspect **test9c_headless_output.txt**.

### Debug output and quiet mode

- **GUI usage (default):** The macro is chatty in the Report view when used interactively. `selection_set_core.DEBUG_SELECTIONSET` and `SELECTIONSET_VERBOSE_CALLBACK` default to **True** so you see what the SelectionSet tools are doing.
- **Automated tests / scripted runs:** The test runner (`run_selectionset_tests`) temporarily sets `SELECTIONSET_VERBOSE_CALLBACK = False` to keep logs readable and avoid per-face diagnostics for every callback. Normal debug prints in `selection_set_core` stay under `DEBUG_SELECTIONSET`, which you can set to **False** in Python if you prefer a quiet session.

**Where to disable debugging:**

| What to disable | Where | What to set |
|-----------------|--------|-------------|
| Observer / view debug (addSelection, attach, etc.) | `selection_set_core.py` (top) | `DEBUG_SELECTIONSET = False` |
| Create/Update SelectionSet callback messages ([DEBUG], per-face/per-solid) | `selection_set_core.py` (top) | `SELECTIONSET_VERBOSE_CALLBACK = False` |
| FEM cantilever & GMSH debug files and "[FEM-cut]" / "[GMSH-mesh]" lines | `selection_set_tests.py` (top) | `WRITE_FEM_DEBUG_FILES = False` (default) |
| SelectionSet debug export after startup tests | `run_macro_and_tests_at_startup.py` | Set env `SELECTIONSET_SKIP_DEBUG_EXPORT=1` before running, or remove/comment out the `export_selectionset_debug_properties(...)` calls |

Debug output is written to **.log** files (e.g. `fem_cantilever_debug.log`, `gmsh_mesh_debug.log`, `selection_set_debug_export.log`) so they are clearly log output; test results stay in `selection_set_test_results.txt`.

---

## Features (reference)

- **Toolbar** – `Selection_tools` toolbar: **Create/Update SelectionSet**, Expand SelectionSet, **Create SelectionSet link**, Build test geometry, **Run Test 1**, **Run Test 3**, Run SelectionSet Tests, **FEM beam example**. Observer is detached automatically on macro run to avoid duplicates.
- **Create/Update SelectionSet** – Select only SelectionSet(s) and click to recompute from stored Main/Shape; or select main + shape (and optionally a solid-filter set) to create or update SelectionSet(s).
- **Volume-based selection** – Main object's faces/solids are filtered by the selection-defining shape: either **fully inside** or **intersection** (for faces); for solids, **fully_inside** uses volume containment (boolean common), and **intersection** selects solids that have any volume overlap with the selection shape.
- **Selection by coordinates** – VolumeMode **coordinates** finds faces or solids that contain a given point. Set **CoordinatePoint** (x,y,z) in the Data tab, or leave it at (0,0,0) and set **SelectionShape** to use that shape’s center. **Create/Update SelectionSet** (with the set selected) recomputes the element list. Create with one object (main) then set CoordinatePoint and Update; or select main + a shape to use the shape’s center as the point.
- **Stored shapes** – MainObject, SelectionShape, Mode (faces/solids), VolumeMode (fully_inside/intersection/coordinates), and optional **CoordinatePoint** (for coordinates mode) are stored on the SelectionSet and editable in the Data tab; **Create/Update SelectionSet** (with the set selected) recomputes the element list from them. The Data tab hides fields that are not relevant (e.g. **CoordinatePoint** when not in coordinates mode, **SolidFilterRefs** when in solids mode) to keep the **Selection** group compact.
- **Filter faces by solid** – For face SelectionSets you can restrict to faces that belong to specific solids (e.g. one half-sphere): select **main object**, **selection shape**, then a **SelectionSet whose ElementList is solids** (e.g. from a prior “Solids” SelectionSet). Or set **SolidFilterRefs** in the Data tab (click **...** to select solids, same as Main Object / Selection Shape).
- **Expand selection** – Double-click or right-click a SelectionSet in the tree (or use **Expand SelectionSet**) to restore its elements in the GUI selection and 3D view; faces are highlighted (solids highlighting not fully supported in FreeCAD 1.0.2). Right-click also shows a **SelectionSet** submenu with actions like **Update SelectionSet** and **Delete SelectionSet only (keep shapes)**.
- **Geometric primitives** – Any Part object (box, sphere, cylinder, BooleanFragments, PartDesign Body, etc.) can be used as main object or selection-defining shape.
- **SelectionSet link** – Tree object that **adds** elements from SelectionSets and/or **direct shapes**, **subtracts** SelectionSets and/or shapes, then applies the result to a **target** (e.g. FEM constraint). Create via **Create SelectionSet link**. In the **Data tab**: **Target objects** (click … to add one or more); **Add selection sets** (SelectionSet objects); **Add shapes** (Part/geometry objects – all faces or solids added, see **Shape ref mode**); **Subtract selection sets**; **Subtract shapes**; **Shape ref mode** = faces | solids (for direct shapes). Selection sets and shapes are evaluated together. The link’s **name (Label)** and **RefSummary** property show a hint of how many and which type of elements are in the combined refs (e.g. “6 faces”, “2 solids, 4 faces”); they are updated when refs are computed (e.g. on **Renew elements**). Right-click the link → **SelectionSet link** submenu: **Renew elements** (apply to target), **Expand to elements (replace selection)**, **Add elements from list to selection**, **Select only elements**. Option **UseOneByOneWorkaround** for constraints that only accept one face at a time. **Updating the target:** The target (FEM constraint, etc.) is updated only when you choose **Renew elements** (right‑click link → SelectionSet link). The link does not apply to the target on property change or document recompute, so FEM objects are not repeatedly touched and you avoid “still touched after recompute” warnings. After changing the link’s inputs or a linked SelectionSet’s content, use **Renew elements** to push refs to the target.
- **FEM material geometry** – In FreeCAD’s geometry reference selector for FEM material (and similar), use **Face, Edge**; **Solid** does not work (tested with FEM material). Use a face-based SelectionSet for material geometry via a link, or select faces/edges manually in the selector.
- **FEM integration** – Use SelectionSets as references for FEM constraints/loads: `get_fem_references_from_selectionset(set_obj)` returns `[(obj, subname), ...]`; `get_fem_references_from_selectionsets([set1, set2, ...])` combines refs from multiple sets (order preserved). `apply_selectionset_to_fem_constraint(set_obj, constraint)` sets References in one go. For constraints that **only accept one face at a time** (e.g. Fem::ConstraintElectrostaticPotential), use `apply_references_to_fem_constraint_one_by_one(refs, constraint, doc)` to add each reference separately (workaround documented in `selection_set_core.py`). Double-clicking a SelectionSet with a FEM constraint selected also transfers refs. **Test FEM** creates an analysis, ConstraintFixed, ConstraintElectrostaticPotential, FEM material (solid), and **three SelectionSetLinks** (fixed constraint, electrostatic constraint, material with solid from Test9b); SKIP if FEM workbench not loaded.
- **FEM beam example** – Toolbar **FEM beam example** builds a cantilever beam (box, optional cylinder cut), SelectionSets for fixed/force face and solid, FEM_Beam_Config (Suppressed = box vs cut), and **FEM_Beam_Variables** / **FEM_Beam_VarSet** for Suppressed and **CylinderRadius** (default 300 mm; edit there—config and cylinder follow). **FEM_Beam_Readme** (App::TextDocument) lists expected deflections. Automated test **Test FEM cantilever full** runs with hole r=300, without hole, and with hole r=400; on 1.0.2 GMSH with hole may fail (see Report view info).

**Note:** A `SelectionSet` itself does **not** expose a `References` property and is not meant to be linked directly as the target of a FEM constraint/material. Instead, use the helper functions above or a `SelectionSetLink` to push its elements into the FEM object’s `References`. This avoids side effects on the SelectionSet’s own behaviour and keeps FEM integration explicit.

#### Gmsh meshes (`Fem::FemMeshGmsh`) and SelectionSetLink

FreeCAD’s Gmsh mesh object (`Fem::FemMeshGmsh`) behaves slightly differently from FEM constraints/materials:

- **Constraints/materials** (e.g. `Fem::ConstraintFixed`, `Fem::Material`) use a `References` property listing `(object, "Face1")`, `(object, "Edge2")`, etc. SelectionSetLink applies to this `References` list.
- **Gmsh mesh** objects do not use `References` for geometry; they refer to a **single shape** via properties such as `Part`, `Geometry`, or `Shape` (depending on FreeCAD version and how the mesh was created). The macro follows the built-in examples (like `ccx_cantilever_base_solid.py`) and assigns **`mesh.Shape`** (or `mesh.Part` / `mesh.Geometry` when present) from the chosen beam solid.

Implications for SelectionSetLink:

- When you want to drive a **Gmsh mesh** from a SelectionSet, target the **mesh’s shape**, not the mesh object’s `References`. In practice this means:
  - Use SelectionSetLinks primarily for **constraints and materials** (targets with a `References` property).
  - For the mesh itself, the macro sets the mesh’s `Shape` from the beam (Beam_Box or `Cut`) and regenerates the Gmsh mesh; there is no separate per-face/solid `References` list on the mesh.
- This mirrors how the built‑in FEM examples work today: mesh generation takes a single solid and produces a volume mesh; all face/edge selections happen on the geometry or on downstream FEM objects, not on the mesh itself.

Suggestions for FreeCAD (future improvement):

- Provide a unified way for FEM elements (constraints, materials, meshes) to expose **shape-based references** (faces/solids) in a consistent property, so SelectionSetLink could target meshes more naturally.
- Consider adding a dedicated “mesh geometry” reference type for Gmsh meshes that aligns with how constraints/materials use `References`, or allow meshes to accept `References` similar to constraints.

- **Test runner** – Built-in test suite (tests 1–9, 9a, 9b, 9c, 9d, 10, 10b, 10c–10f, Test FEM, Test FEM cantilever full) and test geometry builders (two cubes, cube with inner sphere, cylinder with two half-spheres + Body003); use **Run SelectionSet Tests** from the toolbar.
- **Automated test run** – With the AppImage in this folder, `./run_with_gui.sh` starts FreeCAD and runs the macro and full test suite at startup; results are written to **selection_set_test_results.txt**. Use `./run_with_gui.sh --exit` to close FreeCAD after tests (e.g. for scripted runs and reading the results file).
- **Naming and storage** – SelectionSet objects can be named and hold ElementList (faces/solids), Mode, VolumeMode, and optional SolidFilterRefs. New SelectionSets get tree label **SelectionSet …** (e.g. SelectionSet 0, SelectionSet Test1_…); new SelectionSetLinks get **SelectionSetLink …** with ref summary (e.g. SelectionSetLink 1 (6 faces)).

### FEM beam geometry switch (research and implementation)

**Can the switch be done with existing FreeCAD elements only?** No. The geometry switch (use either Beam_Box or Beam_Cut for the same mesh, SelectionSets, and constraints) cannot be achieved with standard FreeCAD objects alone because:

1. **PropertyLink does not support conditional expressions** – The mesh object’s Part/Shape and each SelectionSet’s MainObject are `PropertyLink`. FreeCAD’s expression engine does not allow a link to be set conditionally to one of two objects (e.g. “if VarSet.UseCut then Beam_Cut else Beam_Box”). So we cannot drive “which geometry” by a variable via expressions only.
2. **No built-in object switch** – There is no standard object that takes a boolean and two shapes and outputs one of them as the “active” shape for downstream links.
3. **Imperative update steps** – After changing the geometry reference, we must (a) recompute each SelectionSet from its (new) MainObject to refresh ElementList, and (b) re-apply each SelectionSetLink’s refs to its target (FEM constraints, material). These are imperative operations (Python: `recomputeSelectionSetFromShapes`, `_apply_link_to_target`); they have no expression equivalent.

**How it was implemented:** A custom **App::FeaturePython** object **FEM_Beam_Config** uses a **Suppressed** property (bool): **Suppressed=True** → use box geometry (no hole), **Suppressed=False** → use cut geometry. If the **Beam_Cut** (Part::Cut) object has a native **Suppressed** property (e.g. in FreeCAD 1.1), it is kept in sync so the cut can also be toggled from the tree. Optionally a variable container is created so **Suppressed** can be driven from one place: in **FreeCAD 1.1** the macro uses **App::Variables** (FEM_Beam_Variables) when available; in **1.0** or if Variables is not available it uses **App::VarSet** (FEM_Beam_VarSet). **FEM_Beam_Config.Suppressed** is linked by expression to that object’s **Suppressed** property. **Cylinder radius:** The hole radius is stored as **CylinderRadius** (mm); **default is 300** (never 0, to avoid “Radius of cylinder too small”). When the variable container exists, **CylinderRadius** is added there and is the **source of truth**—edit it in **FEM_Beam_Variables** or **FEM_Beam_VarSet**; the config and **Beam_CylinderCut.Radius** follow by expression. When **Suppressed** changes, the config’s **onChanged** runs (deferred via `QTimer.singleShot` to avoid recursive recompute) and: (1) syncs **Beam_Cut.Suppressed** when present; (2) sets the mesh object’s Part/Shape to `Beam_Box` or `Beam_Cut`; (3) sets **MainObject** on the three SelectionSets and calls **recomputeSelectionSetFromShapes**; (4) applies the three SelectionSetLinks and runs **doc.recompute()**.

## Icons

- **SelectionSet** and **SelectionSetLink** use custom SVG icons from the `icons/` folder so they show in the model tree when the macro is run (including via `run_macro_and_tests_at_startup.FCMacro` or command line). Bundled files: `icons/SelectionSet.svg`, `icons/SelectionSetLink.svg`. To override, replace these with your own SVGs (e.g. 32×32 viewBox).
- **Toolbar buttons** use the same SVGs (icon beside text): Create/Update SelectionSet, Expand, Update, Create SelectionSet link, and the test buttons (Run SelectionSet Tests, Build test geometry, etc.) show the SelectionSet or SelectionSetLink icon next to the label.

## Requirements

- FreeCAD (tested with 1.0.2 and 1.1)
- PySide (Qt), Part, PartDesign (for test geometry)

## Releasing for Addon Manager

See **RELEASE.md** for how to publish this macro so users can install it via FreeCAD’s Addon Manager (official docs links, three submission options, and steps for releasing via your own GitHub repo with `package.xml`).

## Remaining work

See **TODO.md** for the current list of open tasks (tests, documentation, and optional features). This README does not duplicate that list to avoid going out of sync.
