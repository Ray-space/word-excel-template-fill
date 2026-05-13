# word-excel-template-fill

将考务类 **Word（`.docx`）** 试题解析并按 **Excel 模板表头** 写入 **`.xlsx`**，支持题型识别、多选答案分隔符与导出校验摘要。

## 快速开始（Python CLI）

```bash
pip install -r requirements.txt
python word_to_questionbank_excel.py --input 试题.docx --template 模板.xlsx --output 导出.xlsx --module 默认标签 --answer-separator 、
```

Windows 可在项目根目录打开终端，直接执行上节 `python` 命令（需已安装 Python 并将 `python` 加入 PATH）。

打包为单文件 EXE：`powershell -ExecutionPolicy Bypass -File .\build_tool.ps1`，再使用根目录 `run_word_to_excel.bat`（依赖 `dist\word_to_excel.exe`）。

## Web 前端（Next.js）

与 CLI **同一套** Python：`POST /api/convert` 调用 `word_to_questionbank_excel.py`。

```bash
npm install
npm run dev
```

浏览器打开 **http://localhost:3000**（若端口被占用，以终端输出的 Local URL 为准）。上传 Word + Excel 模板即可导出下载。

生产：`npm run build` && `npm run start`。

## Cursor Skill

说明见 [`.cursor/skills/word-excel-template-fill/SKILL.md`](.cursor/skills/word-excel-template-fill/SKILL.md)。

## 仓库

GitHub: <https://github.com/Ray-space/word-excel-template-fill>
