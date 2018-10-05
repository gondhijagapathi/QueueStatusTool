import cx_Freeze
import sys

base = None

if sys.platform == 'win32':
    base = "Win32GUI"

executables = [cx_Freeze.Executable("main.py", base=base, icon="vf_logo.png")]

cx_Freeze.setup(
    name = "SeaofBTC-Client",
    options = {"build_exe": {"packages":["tkinter","matplotlib"], "include_files":["vf_logo.png","excel.png","output_image.png"]}},
    version = "0.01",
    description = "lajlkdslj",
    executables = executables
    )