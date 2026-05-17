@echo off
cd /d "%~dp0"

if not exist ".venv\Scripts\activate.bat" (
    echo Virtual environment not found.
    echo Please create it first with: python -m venv .venv
    pause
    exit /b 1
)

call .\.venv\Scripts\activate.bat

python generate.py

pause
