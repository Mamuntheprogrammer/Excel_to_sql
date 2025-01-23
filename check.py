import sys
from cx_Freeze import setup, Executable
import os

# <added>
import os.path
PYTHON_INSTALL_DIR = os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tcl8.6')
os.environ['TK_LIBRARY'] = os.path.join(PYTHON_INSTALL_DIR, 'tcl', 'tk8.6')
# </added>

base = None
if sys.platform == 'win32':
    base = 'Win32GUI'

# Increase recursion limit
sys.setrecursionlimit(3000)


executables = [
    Executable('ExcelSQL_Studio.py', shortcutName="ExcelSQL Studio",
            icon="logo.ico",base=base)
]


# Now create the table dictionary
shortcut_table = [
    ("DesktopShortcut",        # Shortcut
     "DesktopFolder",          # Directory_
     "ExcelSQL Studi",           # Name
     "TARGETDIR",              # Component_
     "[TARGETDIR]ExcelSQL_Studio.exe",# Target
     None,                     # Arguments
     None,                     # Description
     None,                     # Hotkey
     None,                     # Icon
     None,                     # IconIndex
     None,                     # ShowCmd
     'TARGETDIR'               # WkDir
     )
    ]

# Now create the table dictionary
msi_data = {"Shortcut": shortcut_table}
bdist_msi_options = {'data': msi_data}

# <added>
options = {
    
          "bdist_msi": bdist_msi_options,

    'build_exe': {
        'include_files':[
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tk86t.dll'),
            os.path.join(PYTHON_INSTALL_DIR, 'DLLs', 'tcl86t.dll'),
        'ficon.png','logo.ico',

         ],
         
        
    },
}
# </added>


setup(name = 'ExcelSQL Studio',
      version = "1.0",
      description = 'ExcelSQL Studio Installer',
      # <added>
      author = 'MD : ABDULLAH AL MAMUN',
      author_email = 'pygemsbd@gmail.com',
      options = options,
      # </added>
      executables = executables
      )