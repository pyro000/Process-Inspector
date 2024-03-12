import sys
from cx_Freeze import setup, Executable

base = "Win32GUI" if sys.platform == "win32" else None
base_c = None

includes = []
excludes = []
packages = []
includefiles = ['lib/']
build_exe_options = {'build_exe': {'include_files': includefiles, 'includes': includes}}

exe = Executable("main.py", target_name='MinerInspector.exe', base=base, icon='lib/icon.ico')

setup(
    name='MinerInspector',
    version="1.0",
    description='MinerInspector',
    options=build_exe_options,
    executables=[exe]
)
