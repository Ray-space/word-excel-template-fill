# Word Excel Template Fill 发布说明

## 1) GitHub 仓库建议保留的文件

- `README.md`
- `requirements.txt`
- `word_to_questionbank_excel.py`
- `parse_exam_questions.py`
- `test_parse_exam_questions.py`
- `build_tool.ps1`
- `run_word_to_excel.bat`
- `.cursor/skills/word-excel-template-fill/SKILL.md`

## 2) 不建议提交的内容

- `.pydeps/`
- `.tmp/`
- `build/`
- `dist/`
- `release/`
- `*.spec`
- 导出的结果文件，如 `result.xlsx`

这些内容已经写入 `.gitignore`。

## 3) 打包 EXE

在仓库根目录执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build_tool.ps1
```

成功后会生成：

- `dist\word_to_excel.exe`

## 4) 运行导出

```bat
run_word_to_excel.bat "<input.docx>" "<template.xlsx>" "<output.xlsx>" "<module>" "、"
```

示例：

```bat
run_word_to_excel.bat "E:\题目\2025年企业所得税强基题库255题.docx" "E:\题目\导题模板.xlsx" "E:\题目\result.xlsx" "企业所得税" "、"
```

## 5) 参数说明
- `input.docx`: 题库 Word 文件
- `template.xlsx`: 模板文件（读取首行表头）
- `output.xlsx`: 导出文件
- `module`: 标签列写入值
- `answer_separator`: 多选答案分隔符，可选 `、`（默认）/ `,` / `，` / 空字符串

## 6) Python 入口脚本
- 推荐入口：`word_to_questionbank_excel.py`
- 核心逻辑：`parse_exam_questions.py`

## 7) Skill 安装方式

如果别人是为了安装 Skill，而不是直接运行脚本，只需要复制：

- `.cursor/skills/word-excel-template-fill`

复制到自己的项目 `.cursor/skills/` 或个人 skills 目录即可。
