@echo off
setlocal
set EXE=%~dp0dist\word_to_excel.exe

if not exist "%EXE%" (
  echo [ERROR] 未找到 %EXE%
  echo 请先运行 build_tool.ps1 打包。
  pause
  exit /b 1
)

if "%~4"=="" (
  echo 用法:
  echo   run_word_to_excel.bat ^<input.docx^> ^<template.xlsx^> ^<output.xlsx^> ^<module^> [answer_separator]
  echo 示例:
  echo   run_word_to_excel.bat "C:\path\questions.docx" "C:\path\template.xlsx" "C:\path\result.xlsx" "模块名称" "、"
  pause
  exit /b 1
)

set SEP=%~5
if "%SEP%"=="" set SEP=、

"%EXE%" --input "%~1" --template "%~2" --output "%~3" --module "%~4" --answer-separator "%SEP%"
if errorlevel 1 (
  echo [ERROR] 执行失败
  pause
  exit /b 1
)

echo [OK] 导出完成: %~3
endlocal
