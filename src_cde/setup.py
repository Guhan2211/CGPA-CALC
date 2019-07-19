import os
import sys

PYTHON_INSTALL_DIR=os.path.dirname(os.path.dirname(os.__file__))
os.environ['TCL_LIBRARY']= r'C:\Users\NIRMAL\AppData\Local\Programs\Python\Python36\tcl\tcl8.6'
os.environ['TK_LIBRARY']= r'C:\Users\NIRMAL\AppData\Local\Programs\Python\Python36\tcl\tk8.6'
from cx_Freeze import setup, Executable
# Dependencies are automatically detected, but it might need fine tuning.
build_exe_options = {"packages": ["os","tkinter","numpy","matplotlib"], "excludes": ["turtle"],
                     'include_files':[
                         os.path.join(PYTHON_INSTALL_DIR,'DLLs','tk86t.dll'),
                         os.path.join(PYTHON_INSTALL_DIR,'DLLs','tcl86t.dll')
                         ]
                     }
# GUI applications require a different base on Windows (the default is for a
# console application).
includefiles=["C:\\Users\\NIRMAL\\Desktop\\v5\\icofile.ico"]

base = None

if sys.platform == "win32":
    base = "Win32GUI"
setup( name = "Result Analyzer",
    version = "5.0",
    description = "Analyzes Semester Result!",
    author="Guhan",   
    options = {"build_exe": build_exe_options},
    executables = [Executable("mainscr.py", base=base,icon='C:\\Users\\NIRMAL\\Desktop\\v5\\icofile.ico')])
