@echo off
REM --- SETTINGS
SET "PYTHON_EXE=D:\Programs\Python\Python 314\python.exe"
SET "PROJECT_DIR=D:\AmmashA.IA\database_management"
SET "ERROR_LOG=%PROJECT_DIR%\code_files\utilities\ui\query_UI.py"


START "Running Query Interface" /WAIT /D "%PROJECT_DIR%" cmd.exe /C ^
    ""%PYTHON_EXE%" "%PROJECT_DIR%\code_files\utilities\ui\query_UI.py"

EXIT