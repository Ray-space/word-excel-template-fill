@echo off
chcp 65001 >nul
cd /d "%~dp0"
python word_to_questionbank_excel.py %*
if errorlevel 1 exit /b 1
