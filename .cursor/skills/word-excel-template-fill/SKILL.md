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
   - `word_to_questionbank_excel.py --input <docx> --template <xlsx> --output <xlsx> --module <tag>`
3. If needed, set answer delimiter:
   - `--answer-separator "、"` (default), or `","`, `"，"`, `""`.
4. Report:
   - output path
   - total question count
   - validation summary (`critical`, `warnings`, `pass_rate`)

## Packaged Tool
- Build once: `powershell -ExecutionPolicy Bypass -File .\build_tool.ps1`
- Run tool: `run_word_to_excel.bat <input.docx> <template.xlsx> <output.xlsx> <module> [answer_separator]`

## Guardrails
- Do not overwrite the original template file unless user explicitly asks.
- Keep template header order as output order.
- If mandatory columns are not recognized, surface warnings and request template confirmation.
