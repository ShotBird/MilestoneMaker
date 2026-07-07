@echo off
REM ============================================================
REM  Build "Milestone Maker(v1.0).exe"   (Windows only)
REM  PyInstaller is NOT a cross-compiler -> run this on Windows.
REM  When you bump the version: change VER below AND
REM  VERSION in main.py so the exe name and the window match.
REM ============================================================
setlocal
cd /d %~dp0
set VER=v1.0
set NAME=Milestone Maker(%VER%)

call .venv\Scripts\activate 2>nul

python -c "import tkinter" 2>nul || (echo [ERROR] tkinter missing -- use a python.org Python ^(Tcl/Tk bundled^). & exit /b 1)

pyinstaller --noconfirm --clean --onefile --windowed ^
  --name "%NAME%" ^
  --icon "assets\app.ico" ^
  --add-data "assets\output.pptx;." ^
  --add-data "assets\template_input.xlsm;." ^
  --add-data "assets\app.ico;." ^
  --hidden-import lxml._elementpath ^
  --collect-submodules comtypes ^
  --collect-submodules pptx ^
  --collect-submodules openpyxl ^
  main.py

if errorlevel 1 ( echo BUILD FAILED & exit /b 1 )
echo Built: dist\%NAME%.exe
