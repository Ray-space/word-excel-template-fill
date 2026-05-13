@echo off
chcp 65001 >nul
setlocal
set EXE=%~dp0dist\word_to_excel.exe
set INTERACTIVE=0

if not exist "%EXE%" (
  echo [ERROR] EXE not found: %EXE%
  echo Please make sure dist\word_to_excel.exe exists.
  pause
  exit /b 1
)

if "%~4"=="" (
  set INTERACTIVE=1
  echo [INFO] No command args detected. Interactive mode.
  echo.
  set /p ARG_INPUT=Input docx path: 
  set /p ARG_TEMPLATE=Template xlsx path: 
  set /p ARG_OUTPUT=Output xlsx path: 
  set /p ARG_MODULE=Module/tag text: 
  set /p ARG_SEP=Answer separator (default: 、): 
  if "%ARG_INPUT%"=="" goto :show_usage
  if "%ARG_TEMPLATE%"=="" goto :show_usage
  if "%ARG_OUTPUT%"=="" goto :show_usage
  if "%ARG_MODULE%"=="" goto :show_usage
  if "%ARG_SEP%"=="" set ARG_SEP=、
  goto :run
)

set SEP=%~5
if "%SEP%"=="" set SEP=、

set ARG_INPUT=%~1
set ARG_TEMPLATE=%~2
set ARG_OUTPUT=%~3
set ARG_MODULE=%~4
set ARG_SEP=%SEP%

:run
"%EXE%" --input "%ARG_INPUT%" --template "%ARG_TEMPLATE%" --output "%ARG_OUTPUT%" --module "%ARG_MODULE%" --answer-separator "%ARG_SEP%"
if errorlevel 1 (
  echo [ERROR] 执行失败
  pause
  exit /b 1
)

echo [OK] Export complete: %ARG_OUTPUT%
if "%INTERACTIVE%"=="1" pause
endlocal
exit /b 0

:show_usage
echo.
echo Usage:
echo   run_word_to_excel.bat ^<input.docx^> ^<template.xlsx^> ^<output.xlsx^> ^<module^> [answer_separator]
echo Example:
echo   run_word_to_excel.bat "C:\path\questions.docx" "C:\path\template.xlsx" "C:\path\result.xlsx" "模块名称" "、"
pause
endlocal
exit /b 1
