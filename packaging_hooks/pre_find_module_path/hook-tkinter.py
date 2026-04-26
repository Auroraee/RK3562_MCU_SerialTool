def pre_find_module_path(hook_api):
    # The local Tcl probe can fail even when the tkinter package is importable.
    # Keep tkinter discoverable; hook-_tkinter.py collects the Tcl/Tk data files.
    return
