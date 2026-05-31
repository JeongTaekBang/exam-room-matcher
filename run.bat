@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "CONDA_BAT="
if exist "%USERPROFILE%\miniconda3\condabin\conda.bat"  set "CONDA_BAT=%USERPROFILE%\miniconda3\condabin\conda.bat"
if exist "%USERPROFILE%\Miniconda3\condabin\conda.bat"  set "CONDA_BAT=%USERPROFILE%\Miniconda3\condabin\conda.bat"
if exist "%USERPROFILE%\anaconda3\condabin\conda.bat"   set "CONDA_BAT=%USERPROFILE%\anaconda3\condabin\conda.bat"
if exist "%USERPROFILE%\Anaconda3\condabin\conda.bat"   set "CONDA_BAT=%USERPROFILE%\Anaconda3\condabin\conda.bat"

if defined CONDA_BAT (
    call "%CONDA_BAT%" activate base
    python -X utf8 -m streamlit run dashboard.py
) else (
    python -X utf8 -m streamlit run dashboard.py
)
pause
