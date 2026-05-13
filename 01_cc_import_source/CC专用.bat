@echo off
setlocal
chcp 65001 >nul

set "SCRIPT_DIR=%~dp0"
set "ENTRY=%SCRIPT_DIR%CC专用.py"
set "SINGLE_ENTRY=%SCRIPT_DIR%word_to_questionbank_excel.py"
set "CORE_PARSER=%SCRIPT_DIR%parse_exam_questions.py"
set "PY_CMD="

echo [INFO] 正在检查运行环境...

where py >nul 2>nul
if not errorlevel 1 (
  set "PY_CMD=py -3"
)

if "%PY_CMD%"=="" (
  where python >nul 2>nul
  if not errorlevel 1 (
    set "PY_CMD=python"
  )
)

if "%PY_CMD%"=="" (
  echo [ERROR] 未检测到 Python 命令。
  echo [建议] 请先安装 Python 3.10+，并勾选 "Add python to PATH"。
  pause
  exit /b 1
)

if not exist "%ENTRY%" (
  echo [ERROR] 缺少入口脚本: %ENTRY%
  echo [建议] 请确认你解压的是完整 A 包。
  pause
  exit /b 1
)

if not exist "%SINGLE_ENTRY%" (
  echo [ERROR] 缺少导题脚本: %SINGLE_ENTRY%
  echo [建议] 请确认同目录包含 word_to_questionbank_excel.py。
  pause
  exit /b 1
)

if not exist "%CORE_PARSER%" (
  echo [ERROR] 缺少解析脚本: %CORE_PARSER%
  echo [建议] 请确认同目录包含 parse_exam_questions.py。
  pause
  exit /b 1
)

%PY_CMD% -I -c "import openpyxl" >nul 2>nul
if errorlevel 1 (
  echo [ERROR] 缺少依赖: openpyxl
  echo [建议] 在终端执行: pip install openpyxl
  pause
  exit /b 1
)

echo [INFO] 环境检查通过，启动 CC 专用导题...
%PY_CMD% "%ENTRY%"
if errorlevel 1 (
  echo.
  echo [ERROR] 执行失败，请按以下顺序检查：
  echo   1) 模板文件是否存在（企业所得税.xlsx / 增值税.xlsx / 个人所得税.xlsx）
  echo   2) 输入路径是否是存在的 .docx 文件夹
  echo   3) 同目录脚本是否完整（CC专用.py / word_to_questionbank_excel.py / parse_exam_questions.py）
  pause
  exit /b 1
)

echo.
echo [OK] 执行完成。
pause
endlocal
