@echo off
REM --- SETTINGS
SET "PYTHON_EXE=D:\Programs\Python\Python 314\python.exe"
SET "PROJECT_DIR=D:\AmmashA.IA\OneDrive - University of Leeds\UOL Data Manager\database_management"

ECHO Starting Cohort Mapping Update Job...

START "Running Cohort Mapping Update Job" /WAIT /D "%PROJECT_DIR%" cmd.exe /C ^
    ""%PYTHON_EXE%" "%PROJECT_DIR%\code_files\automation\update_cohort_mapping.py"

ECHO Job finished at %TIME%.
TIMEOUT /T 5 /NOBREAK
EXIT