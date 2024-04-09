from cx_Freeze import setup, Executable

base = None
executables = [Executable("app.py", base=base, target_name="WorkTimeApp.exe")]

packages = ["tkinter", "tkinterdnd2", "pandas", "datetime"]
options = {
    'build_exe': {
        'packages': packages,
        'include_files': [],
    },
}

setup(
    name="WorkTimeApp",
    options=options,
    version="1.0",
    description='Work Time Rounding Tool',
    executables=executables
)