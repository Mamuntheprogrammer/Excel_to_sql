import sys
from cx_Freeze import setup, Executable
import os

# Set up paths for TCL and TK libraries
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')

# Base for Windows GUI
base = None
if sys.platform == 'win32':
    base = 'Win32GUI'


# Increase recursion limit
sys.setrecursionlimit(3000)


# Executable definition
executables = [
    Executable(
        'ExcelSQL_Studio.py',
        icon="logo.ico",
        base=base,
    )
]

# MSI Shortcut table
shortcut_table = [
    (
        "DesktopShortcut",  # Shortcut
        "DesktopFolder",  # Directory_
        "ExcelSQL Studio",  # Name
        "TARGETDIR",  # Component_
        "[TARGETDIR]ExcelSQL_Studio.exe",  # Target
        None,  # Arguments
        "Launch ExcelSQL Studio",  # Description
        None,  # Hotkey
        "logo.ico",  # Icon
        None,  # IconIndex
        None,  # ShowCmd
        "TARGETDIR",  # WkDir

        
    )
]

# MSI Options
msi_data = {"Shortcut": shortcut_table}
bdist_msi_options = {
    "data": msi_data,
    "add_to_path": False,
}

# Build Options
options = {
    "bdist_msi": bdist_msi_options,
    "build_exe": {
        "include_files": [
            os.path.join(PYTHON_INSTALL_DIR, "DLLs", "tk86t.dll"),
            os.path.join(PYTHON_INSTALL_DIR, "DLLs", "tcl86t.dll"),
            "logo.ico",  # Application icon
            "ficon.png",
        ],
        "packages": [
            "tkinter",
            "pandas",
            "pandastable",
            "sqlalchemy",
            "tkinterdnd2",
            "base64",
        ],
    },
}

# Setup definition
setup(
    name="ExcelSQL_Studio",
    version="1.0",
    description="Run SQL Queries on Excel Files",
    author="MD: Abdullah Al Mamun",
    author_email="pygemsbd@gmail.com",
    options=options,
    executables=executables,
)
