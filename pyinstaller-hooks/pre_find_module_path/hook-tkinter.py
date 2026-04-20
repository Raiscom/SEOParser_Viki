"""Override PyInstaller's tkinter exclusion for the bundled runtime."""


def pre_find_module_path(hook_api):
    """Keep tkinter importable even if the source runtime lacks working Tcl discovery."""
    return
