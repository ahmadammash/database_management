@echo off
REM --- SETTINGS
SET "PYTHON_EXE=D:\Programs\Python\Python 314\python.exe"
SET "PROJECT_DIR=D:\AmmashA.IA\OneDrive - University of Leeds\UOL Data Manager\database_management"
SET "ERROR_LOG=%PROJECT_DIR%\text_files\sync_job_onedrivecnxn_error.log"

REM --- RESET LOG FILE
> "%ERROR_LOG%" (
    ECHO ========================================================================
    ECHO Started: %DATE% %TIME%
    ECHO ========================================================================
)

ECHO Starting Data Synchronization Job...

START "Running Data Sync Job" /WAIT /D "%PROJECT_DIR%" cmd.exe /C ^
    ""%PYTHON_EXE%" "%PROJECT_DIR%\code_files\syncing_scripts\sync_job_cnxn.py" 2>> "%ERROR_LOG%""

ECHO Job finished at %TIME%.
TIMEOUT /T 5 /NOBREAK
EXIT