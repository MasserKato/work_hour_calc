from cx_Freeze import setup, Executable

base = "Win32GUI"
executables = [Executable("app.py", base=base, target_name="WorkTimeApp.exe")]

packages = ["tkinter", "tkinterdnd2", "datetime"]
include_files = []
options = {
    'build_exe': {
        'packages': packages,
        'include_files': include_files,
        # アプリが使わないテスト用・通信用モジュールを同梱しない。
        'excludes': [
            'tcl', 'ttk', 'numpy', 'pandas',
            'test', 'tests', 'unittest', 'pydoc', 'doctest', 'ensurepip', 'venv',
            'sqlite3', 'xmlrpc', 'ftplib', 'poplib', 'imaplib', 'smtplib',
        ],
    },
}

setup(
    name="WorkTimeApp",
    options=options,
    version="1.0",
    description='Work Time Rounding Tool',
    executables=executables
)
