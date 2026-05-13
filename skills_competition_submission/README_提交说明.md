# Skills 比赛提交包说明

本目录已按赛方要求整理，打包为 **ZIP** 根目录应包含下列内容（可直接压缩本文件夹）。

## 1. skill.md（必须）

Skill 核心配置与使用说明，与 Cursor 仓库内 `.cursor/skills/word-excel-template-fill/SKILL.md` 语义一致；本包内文件名为赛方要求的 **`skill.md`**。

## 2. Python 与辅助文件

| 文件 | 说明 |
|------|------|
| `word_to_questionbank_excel.py` | 入口脚本 |
| `parse_exam_questions.py` | 解析与写 Excel 核心逻辑 |
| `requirements.txt` | 运行依赖（主要为 openpyxl） |
| `run_python_cli.bat` | Windows 下用 Python 直接跑入口（评审友好） |
| `build_tool.ps1` | PyInstaller 打包为单文件 EXE |
| `run_word_to_excel.bat` | 调用 `dist\word_to_excel.exe`（需先执行 build） |
| `data/` | 示例与数据说明 |

### 快速运行（Python）

```bash
pip install -r requirements.txt
python word_to_questionbank_excel.py --input 试题.docx --template 模板.xlsx --output 导出.xlsx --module 默认标签 --answer-separator 、
```

## 3. 演示视频链接

编辑 **`演示视频链接.txt`**：将 B 站等平台上的 **完整演示视频（≤3 分钟）** URL 粘贴到「视频链接」一行。

## 4. 商业价值说明书

见 **`商业价值说明书.txt`**（约 200 字）。

## 生成 ZIP（可选）

在项目根目录执行（PowerShell）。下面命令会生成 **内含 `skills_competition_submission` 文件夹** 的压缩包，解压后目录结构清晰：

```powershell
Compress-Archive -Path ".\skills_competition_submission" -DestinationPath ".\word-excel-template-fill_skills-submission.zip" -Force
```

若赛方要求 ZIP **根目录即为文件**（不要外层文件夹），可改用：

```powershell
Compress-Archive -Path ".\skills_competition_submission\*" -DestinationPath ".\word-excel-template-fill_skills-submission-flat.zip" -Force
```

提交前请确认 ZIP 内 **`skill.md`** 文件名与赛方要求完全一致。
