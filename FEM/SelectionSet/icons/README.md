# Custom icons for SelectionSet macro

Bundled SVG icons used in the model tree when the macro is run (including via `run_macro_and_tests_at_startup.FCMacro` or command line).

- **SelectionSet.svg** (or selection_set.svg) – two overlapping squares: one **yellow dashed** (box-selection style) and one **blue solid**.
- **SelectionSetLink.svg** – References **FEM_Analysis.svg** by filename (linked inside the SVG). If **FEM_Analysis.svg** is present in this same folder, it is shown (copy it from your FreeCAD install: `Mod/Fem/Gui/Resources/icons/FEM_Analysis.svg`, or from [wiki](https://wiki.freecad.org/File:FEM_Analysis.svg)). No download script; when the file is missing a simple fallback “A” is shown.

Icons are resolved from the macro `icons/` directory. To override, replace these files with SVGs of the same name (e.g. 32×32 viewBox).
