import sys
from cx_Freeze import setup, Executable

base = None
if (sys.platform == "win32"):
    base = "Win32GUI"

setup(  name = "ImageRenamer",
        version = "0.1",
        options = {"build_exe": {"include_msvcr": True}, "bdist_msi": {"add_to_path": True}},
        executables = [Executable("ImageRenamer.py", base=base)])
