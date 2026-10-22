@echo off
REM --- SETTINGS
SET "PYTHONIOENCODING=utf-8"
SET "PYTHON_EXE=C:\Users\medaamm\AppData\Local\miniconda3\python.exe"
SET "PROJECT_DIR=C:\Users\medaamm\OneDrive - University of Leeds\UOL Data Manager\database_management"

REM --- DEFINE LOG PATH & ENSURE DIRECTORY EXISTS
SET "LOG_DIR=%PROJECT_DIR%\text_files"
SET "ERROR_LOG=%LOG_DIR%\sync_job_error_tristar.log"

IF NOT EXIST "%LOG_DIR%" MKDIR "%LOG_DIR%"

REM --- RESET LOG FILE
(
    ECHO ========================================================================
    ECHO Started: %DATE% %TIME%
    ECHO ========================================================================
) > "%ERROR_LOG%"

ECHO Starting TriStar Server Data Synchronization Job...

START "Running TriStar Server Data Sync Job" /WAIT /D "%PROJECT_DIR%" cmd.exe /C ^
    ""%PYTHON_EXE%" "%PROJECT_DIR%\code_files\syncing_scripts\sync_job.py" 2>> "%ERROR_LOG%""

REM --- ADD FINISH TIME TO LOG
(
    ECHO ========================================================================
    ECHO Finished TriStar Server Data Sync Job: %DATE% %TIME%
    ECHO ========================================================================
) >> "%ERROR_LOG%"

ECHO Job finished at %TIME%.
TIMEOUT /T 5 /NOBREAK
EXIT