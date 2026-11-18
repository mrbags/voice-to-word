@echo off
REM ============================================================================
REM  run.bat - start Voice-to-Word straight from Python (no build needed).
REM
REM  The first time it runs, it creates a private Python environment in a
REM  ".venv" folder and installs the packages from requirements.txt into it,
REM  so nothing is installed system-wide. Later runs skip straight to launching.
REM ============================================================================

REM Work from the folder this script is in, wherever it was started from.
cd /d "%~dp0"

if not exist .venv (
    echo First run: setting up... this takes a minute.
    REM Create the private environment. Fails if Python is not installed or not on PATH.
    python -m venv .venv || (echo Python 3.9+ is required: https://www.python.org/downloads/ & pause & exit /b 1)
    .venv\Scripts\python -m pip install --upgrade pip
    .venv\Scripts\python -m pip install -r requirements.txt || (pause & exit /b 1)
)

REM pythonw runs the app without leaving a black console window open.
start "" .venv\Scripts\pythonw voice_to_word.py
