@echo off
echo Starting JC Tracker Setup...

REM Check for Python 3.9+
python --version 2>NUL | findstr /R "Python 3\.[9-9]\.|Python 3\.[1-9][0-9]" >nul
IF %ERRORLEVEL% NEQ 0 (
    echo Python 3.9+ not found. Please install Python from https://www.python.org/downloads/
    pause
    exit /b 1
) ELSE (
    echo Python 3.9+ found - continuing setup...
)

REM Create virtual environment if it doesn't exist
IF NOT EXIST venv (
    echo Creating virtual environment...
    python -m venv venv
)

REM Activate virtual environment
echo Activating virtual environment...
call venv\Scripts\activate

REM Upgrade pip
echo Upgrading pip...
python -m pip install --upgrade pip

REM Install requirements
echo Installing required packages...
IF EXIST requirements.txt (
    pip install -r requirements.txt
) ELSE (
    echo requirements.txt not found!
    pause
    exit /b 1
)

REM Run Streamlit app
echo Starting JC Tracker...
python -m streamlit run app.py
pause
