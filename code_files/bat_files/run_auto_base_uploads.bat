@echo off
REM --- SETTINGS
SET "PYTHON_EXE=D:\Programs\Python\Python 314\python.exe"
SET "PROJECT_DIR=D:\AmmashA.IA\OneDrive - University of Leeds\UOL Data Manager\database_management"
SET "ERROR_LOG=%PROJECT_DIR%\text_files\auto_base_uploads_error.log


START "Running Auto Base Uploads" /WAIT /D "%PROJECT_DIR%" cmd.exe /C ^
    ""%PYTHON_EXE%" "%PROJECT_DIR%\code_files\syncing_scripts\sync_job_auto_base_uploads.py" 2>> "%ERROR_LOG%"

ECHO Job finished at %TIME%.
TIMEOUT /T 5 /NOBREAK
EXIT