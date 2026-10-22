@echo off
REM --- SETTINGS
SET "PYTHON_EXE=D:\Programs\Python\Python 314\python.exe"
SET "PROJECT_DIR=D:\AmmashA.IA\OneDrive - University of Leeds\UOL Data Manager\database_management"

ECHO Starting Patient List Re-Upload...

START "Running Patient List Re-Upload" /WAIT /D "%PROJECT_DIR%" cmd.exe /C ^
    ""%PYTHON_EXE%" "%PROJECT_DIR%\code_files\syncing_scripts\sync_job_patientList.py"

START "Running TriStar Pathway Re-Upload" /WAIT /D "%PROJECT_DIR%" cmd.exe /C ^
    ""%PYTHON_EXE%" "%PROJECT_DIR%\code_files\base_scripts\processing_tristar_pathway.py"

ECHO Job finished at %TIME%.
TIMEOUT /T 5 /NOBREAK
EXIT