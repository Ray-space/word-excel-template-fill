---
name: word-excel-template-fill
description: Parse exam-style Word documents and fill Excel templates by header semantics. Use when the user asks for Word to Excel conversion, template-aligned field mapping, question type classification, or batch exam import/export.
---

# Word Excel Template Fill

## Purpose
Convert `.docx` exam text into a template-aligned `.xlsx` file with stable column mapping.

## When To Use
- User asks to import Word questions into Excel.
- User provides a template and wants data filled by matching headers.
- User asks to classify question types (single/multiple/judgment) and export in bulk.

## Workflow
1. Ensure input files exist: source `.docx`, template `.xlsx`.
2. Run the parser:
   - `python word_to_questionbank_excel.py --input <docx> --template <xlsx> --output <xlsx> --module <tag>`
3. If needed, set answer delimiter:
   - `--answer-separator "、"` (default), or `","`, `"，"`, `""`.
4. Report:
   - output path
   - total question count
   - validation summary (`critical`, `warnings`, `pass_rate`)

## Packaged Tool
- **Python 环境（推荐评审）**：`pip install -r requirements.txt` 后执行上节命令；Windows 可用 `run_python_cli.bat` 传入相同参数。
- **打包为 EXE**：`powershell -ExecutionPolicy Bypass -File .\build_tool.ps1`，再使用 `run_word_to_excel.bat`（依赖 `dist\word_to_excel.exe`）。

## Web Frontend（Next.js）
仓库根目录 Next 应用与 CLI **同一套** Python：`POST /api/convert` 调 `word_to_questionbank_excel.py`。
1. `npm install` → `npm run dev`
2. 浏览器打开 **http://localhost:3000**（端口以终端为准）
3. 上传 `.docx` + 模板 `.xlsx`，配置分隔符与输出名后导出下载。

## Guardrails
- Do not overwrite the original template file unless user explicitly asks.
- Keep template header order as output order.
- If mandatory columns are not recognized, surface warnings and request template confirmation.
