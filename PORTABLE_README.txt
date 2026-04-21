Word Excel Template Fill - Portable Package

No Python installation is required.

Files:
- run_word_to_excel.bat
- dist\word_to_excel.exe

How to use:
1) Put your input .docx and template .xlsx on your computer.
2) Open Command Prompt in this folder.
3) Run:

run_word_to_excel.bat "<input.docx>" "<template.xlsx>" "<output.xlsx>" "<module>" "、"

Example:
run_word_to_excel.bat "D:\questions.docx" "D:\template.xlsx" "D:\result.xlsx" "企业所得税" "、"

Arguments:
- input.docx: source Word file
- template.xlsx: import template file
- output.xlsx: output Excel path
- module: value written to tag/module column
- answer separator: optional, one of "、" "," "，" or empty
