# FreeCAD SelectionSet macro

SelectionSet creates reusable selections of faces or solids from spatial
regions or coordinates. SelectionSetLink combines and applies those
selections to FEM constraints, materials, or other objects with a
`References` property.

## Installation and first run

1. Copy the complete contents of this folder into a FreeCAD macro directory.
   Keep the `icons/` subfolder beside the Python files.
2. In FreeCAD, choose **Macro → Macros → Execute** and run
   `SelectionSet.FCMacro`.
3. The `Selection_tools` toolbar appears. The normal user actions are:
   - **Create/Update SelectionSet** – create a set from the current main
     object and selection shape, or recompute selected sets from their stored
     properties.
   - **Expand SelectionSet** – restore the stored elements to the GUI
     selection and 3D view.
   - **Create SelectionSet link** – combine sets or direct shapes and apply
     the result to one or more target objects. After changing a link, use
     **SelectionSet link → Renew elements**.

The macro entry point is `SelectionSet.FCMacro`. The files must remain together: `selection_set_core.py`,
`highlight_subelements.py`, `selection_set_tests.py`, the `icons/` folder, and
this README. The test module supplies the optional test toolbar; normal
SelectionSet use does not require running the tests.

## SelectionSet usage

A SelectionSet stores:

- **MainObject** – the object whose faces or solids are selected;
- **SelectionShape** – a box, sphere, body, or other shape defining the region;
- **Mode** – `faces` or `solids`;
- **VolumeMode** – `fully_inside`, `intersection`, or `coordinates`;
- **CoordinatePoint** – used by `coordinates` mode;
- **SolidFilterRefs** – optional solid filters for face selections.

For a basic selection, select the main object and a selection shape, click
**Create/Update SelectionSet**, and choose **Faces** or **Solids**. To refresh
a set after changing its stored properties, select the set and click the same
button.

In `coordinates` mode, set **CoordinatePoint** in the Data tab. Alternatively,
leave it at `(0, 0, 0)` and use the center of the SelectionShape. The
coordinate-point marker can be hidden through the Data or View tab.

To filter faces by a particular solid, set **SolidFilterRefs** through the
`...` button in the Data tab. In the link dialog, select the main object and
enter sub-elements such as `Solid2` or `Solid3` in the second row.

Double-click a SelectionSet to expand its elements. The context menu provides
actions such as **Expand to elements (replace selection)** and **Delete
SelectionSet only (keep shapes)**. Shift-click selects only the main object,
Alt-click selects only the selection shape, and Ctrl-click expands the
elements.

## SelectionSetLink and FEM

A SelectionSetLink can add or subtract SelectionSets and direct shapes, then
apply the combined references to one or more targets. Configure **Target
objects**, **Add selection sets**, **Add shapes**, **Subtract selection sets**,
**Subtract shapes**, and **Shape ref mode** in the Data tab.

Use face-based SelectionSets for FEM constraints and material geometry. FEM
material geometry normally requires **Face, Edge** references rather than a
solid reference. A target is updated only by **Renew elements**.

The FEM beam example, when available in the toolbar, creates a beam, FEM
objects, SelectionSets, and SelectionSetLinks. Its variables include
`Suppressed` and `CylinderRadius`; the default cylinder radius is 300 mm.

## Python API

After running `SelectionSet.FCMacro`, the core module can be used from the
FreeCAD Python console:

```python
import selection_set_core
```

Useful functions include:

- `get_faces_by_point(obj, point, tol=1e-6, solid_filter_list=None)`
- `get_solids_by_point(obj, point, tol=1e-6)`
- `get_faces_by_selection_shape(obj, shape, volume_mode="fully_inside")`
- `create_selectionset(...)`
- `recomputeSelectionSetFromShapes(set_obj)`
- `createSelectionSetLink(...)`
- `get_fem_references_from_selectionset(set_obj)`
- `apply_selectionset_to_fem_constraint(set_obj, constraint)`

For constraints that accept only one face at a time, use
`apply_references_to_fem_constraint_one_by_one(...)`.

## Requirements and limitations

The macro requires FreeCAD with the Part and Qt modules. FEM features require
the FreeCAD FEM workbench. It was tested with FreeCAD 1.0.2, 1.1 RC3,
1.1.4, and 26.3 RC2.

FreeCAD 1.0.2 can fail to generate the GMSH cantilever mesh with a hole.
This is a compatibility limitation of that build; the other tested builds
complete the full test suite.

Generated test tools, AppImages, result files, development plans, and release
metadata are maintained in the development workspace and are not required to
install or use this macro.
