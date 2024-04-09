from cx_Freeze import setup, Executable

base = "Win32GUI"
executables = [Executable("app.py", base=base, target_name="WorkTimeApp.exe")]

packages = ["tkinter", "tkinterdnd2", "pandas", "datetime"]
include_files = []
options = {
    'build_exe': {
        'packages': packages,
        'include_files': include_files,
        'excludes': ['tcl', 'ttk', 'tkinter'],
    },
}

setup(
    name="WorkTimeApp",
    options=options,
    version="1.0",
    description='Work Time Rounding Tool',
    executables=executables
)