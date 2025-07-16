@echo off
title Professor Report Generator
echo Starting Professor Report Generator...

REM Check if virtual environment Python exists locally
if exist "venv\Scripts\python.exe" (
    echo Using virtual environment Python: venv\Scripts\python.exe
    "venv\Scripts\python.exe" run_gui.py
) else (
    echo Virtual environment not found, trying system Python...
    python run_gui.py
)

if errorlevel 1 (
    echo.
    echo Error: Python or required packages not found.
    echo Please ensure the virtual environment is set up correctly.
    echo Or run: pip install -r requirements.txt
    echo.
    pause
)
