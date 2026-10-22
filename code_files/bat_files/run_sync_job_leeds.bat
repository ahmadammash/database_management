@echo off
REM --- FORCE WINDOWS CMD AND PYTHON TO USE UTF-8 ENCODING
chcp 65001 > NUL
set PYTHONIOENCODING=utf-8

REM --- SETTINGS
SET "PYTHON_EXE=D:\Programs\Python\Python 314\python.exe"
SET "PROJECT_DIR=D:\AmmashA.IA\OneDrive - University of Leeds\UOL Data Manager\database_management"
SET "LOG_DIR=%PROJECT_DIR%\text_files"
SET "ERROR_LOG=%LOG_DIR%\sync_job_error_leeds.log"

IF NOT EXIST "%LOG_DIR%" mkdir "%LOG_DIR%"

CD /D "%PROJECT_DIR%"
SET "PYTHONPATH=%PROJECT_DIR%"

(
    ECHO ========================================================================
    ECHO Started: %DATE% %TIME%
    ECHO ========================================================================
) > "%ERROR_LOG%"

ECHO Starting Leeds SQL Server Data Synchronization Job...

REM --- REDIRECT STDERR (2>>) TO ERROR LOG FILE WHILE PRINTING STDOUT TO CONSOLE
"%PYTHON_EXE%" -u "%PROJECT_DIR%\code_files\syncing_scripts\sync_job_combined.py" 2>> "%ERROR_LOG%"

(
    ECHO ========================================================================
    ECHO Finished Leeds SQL Server Data Sync Job: %DATE% %TIME%
    ECHO ========================================================================
) >> "%ERROR_LOG%"

ECHO Job finished at %TIME%.
PAUSE
EXIT