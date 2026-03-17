# -*- coding: utf-8 -*-
"""
FreeCAD SelectionSet macro – entry point.
Run this file as a macro in FreeCAD to load the SelectionSet toolbar and optional test geometry.

Compatibility: FreeCAD 1.0.x and 1.1.x. Core and tests use version-agnostic APIs and
where needed adopt 1.1 properties (e.g. PartDesign Pad SideType) while keeping 1.0 paths.

Modules:
  selection_set_core – SelectionSet logic and GUI (Create/Expand/Update, observer, commands).
  selection_set_tests – Test runner and test geometry builders (Run tests button, build_test_*).
See README.md for usage.
"""
import FreeCAD
import importlib
import os
import sys

# Ensure this macro's directory is on the path so selection_set_core and selection_set_tests can be imported
_macro_dir = os.path.dirname(os.path.abspath(__file__))
if _macro_dir not in sys.path:
    sys.path.insert(0, _macro_dir)

import selection_set_core
import selection_set_tests
# Reload so re-running the macro always uses the latest code from disk
importlib.reload(selection_set_core)
importlib.reload(selection_set_tests)

# Prominent banner so users can see where this macro run starts
_version = getattr(selection_set_core, "__version__", "?")
print("" + "=" * 60)
print("  SelectionSet macro – RUN STARTED  (version %s)" % _version)
print("=" * 60)

# Ensure no stale observer from a previous run (avoids duplicate observers when re-running the macro)
selection_set_core._detach_selectionset_observer(quiet=True)

# Main toolbar: Create/Update SelectionSet, Expand SelectionSet, Create SelectionSet link
selection_set_core.add_selectionset_button_to_toolbar()

# Test toolbar: single "Tests" pulldown (Run Test 1, Build test geometry, Run Test 3, Run all tests, FEM beam example)
selection_set_tests.add_tests_pulldown_to_toolbar()

if selection_set_core.USE_SELECTION_OBSERVER:
    selection_set_core._attach_selectionset_observer()
    print("SelectionSet macro loaded: Observer enabled (single-click expands). Buttons on 'Selection_tools' toolbar.")
else:
    print("SelectionSet macro loaded: Select a SelectionSet in the tree, then click 'Expand SelectionSet' on the toolbar; or double-click/right-click the SelectionSet in the tree.")
print("Next steps: Use the 'Tests' pulldown (Run Test 1, Build test geometry, Run Test 3, Run all tests, FEM beam example).")
