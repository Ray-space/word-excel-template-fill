# Word Excel Template Fill

将试题型 `Word(.docx)` 文档按 `Excel` 模板首行表头语义自动分类并导出为 `.xlsx`，同时附带一个可复用的 Cursor Skill，适合题库批量导入、模板映射和题型分类场景。

开源定位：这是一个面向题库导入场景的轻量工具仓库，目标是用本地、可审计、可重复执行的方式，把结构较松散的 Word 试题整理为符合平台模板的 Excel 数据，并让 Cursor 能通过 Skill 复用同一套流程。

## 项目简介

`Word Excel Template Fill` 解决的是一个常见但很费人工的问题：题库往往保存在 `.docx` 文档里，而目标平台要求按固定的 `.xlsx` 模板上传。这个项目会读取 Word 题目内容，识别题干、选项、答案、解析、难度、知识点等字段，再按模板表头语义写回 Excel。

它同时包含一个 Cursor Skill，适合希望把“Word 转 Excel 导题流程”标准化、沉淀到团队工作流里的场景。

## 为什么适合开源

- 本地运行，不依赖在线模型或外部 API
- 输入输出明确，适合被二次集成或定制
- 同时提供源码、测试、打包脚本和 Skill 定义
- 适合教培、题库运营、内容迁移等重复性导题任务

## 适用场景

- 教培机构运营批量整理题库并上传至题库平台
- 教研团队将 Word 题库沉淀为平台导入模板
- 存量题库迁移（Word/文档格式）到在线题库系统（不限学科）
- 按题型、答案、解析、难度、知识点等字段标准化入库

## 特性

- 本地解析，无需大模型 API，不消耗 token
- 按模板表头自动映射：题型、题干、选项 A-H、答案、题目解析、难度、知识点、标签、指标
- 支持答案分隔符配置（`、` / `,` / `，` / 空）
- 自动处理模板尾部空列，避免 `Unnamed: n` 空白列
- 导出后自带校验摘要（`critical` / `warnings` / `pass_rate`）

## 仓库结构

- `word_to_questionbank_excel.py`：主入口脚本（推荐直接调用）
- `parse_exam_questions.py`：核心解析与导出逻辑
- `test_parse_exam_questions.py`：单元测试
- `build_tool.ps1`：一键打包 `exe`
- `run_word_to_excel.bat`：调用打包工具执行导出
- `.cursor/skills/word-excel-template-fill/SKILL.md`：项目级 Skill

这个仓库已经清理为适合公开发布的最小集合，不包含本地依赖目录、临时脚本、样例产物和构建输出。

## Skill 安装

如果你是把这个仓库作为 Cursor Skill 发布到 GitHub，安装时复制下面这个目录到本地 Cursor skills 目录即可：

- `.cursor/skills/word-excel-template-fill`

推荐两种方式：

1. 直接克隆仓库后，把 `.cursor/skills/word-excel-template-fill` 复制到你自己的项目 `.cursor/skills/` 下。
2. 下载仓库 ZIP，只取出 `word-excel-template-fill` 这个 skill 目录使用。

Skill 主要负责告诉 Cursor：

- 什么时候应该使用这个流程
- 应该调用哪个脚本
- 输出结果时要汇报哪些校验信息

## 环境要求

- Windows + Python 3.10+
- 运行依赖：

```bash
pip install -r requirements.txt
```

若需要打包：

```bash
pip install pyinstaller
```

## 快速开始（Python 运行）

```bash
python word_to_questionbank_excel.py ^
  --input "E:\题目\2025年企业所得税强基题库255题.docx" ^
  --template "E:\题目\导题模板.xlsx" ^
  --output "E:\题目\result.xlsx" ^
  --module "企业所得税" ^
  --answer-separator "、"
```

参数说明：

- `--input`：Word 题库路径（仅支持 `.docx`）
- `--template`：Excel 模板路径（读取首行表头）
- `--output`：导出结果路径
- `--module`：写入“标签”列的值
- `--answer-separator`：多选答案分隔符，默认 `、`

## 打包成可直接运行工具（EXE）

在项目目录执行：

```powershell
powershell -ExecutionPolicy Bypass -File .\build_tool.ps1
```

生成：

- `dist\word_to_excel.exe`

## 用 BAT 调用 EXE

```bat
run_word_to_excel.bat "<input.docx>" "<template.xlsx>" "<output.xlsx>" "<module>" "、"
```

示例：

```bat
run_word_to_excel.bat "E:\题目\2025年企业所得税强基题库255题.docx" "E:\题目\导题模板.xlsx" "E:\题目\result.xlsx" "企业所得税" "、"
```

## GitHub 发布建议

这个仓库现在已经适合直接发到 GitHub：

- 已包含 `README.md`
- 已包含 `requirements.txt`
- 已包含 `.gitignore`
- Skill 文件位于 `.cursor/skills/word-excel-template-fill/SKILL.md`

推荐发布内容：

- 提交源码、测试、打包脚本、README、Skill
- 不提交 `dist/`、`build/`、`*.spec`、`.pydeps/`、`.tmp/`、`release/`、导出结果文件

如果想同时发布可执行版本：

- 在本地运行 `powershell -ExecutionPolicy Bypass -File .\build_tool.ps1`
- 将生成的 `dist/word_to_excel.exe` 上传到 GitHub Releases
- 版本号推荐使用 `v1.0.0`、`v1.0.1` 这种格式

## 质量与校验

运行后会输出：

- 识别题目数量
- 三阶段分流统计（单选/多选/判断）
- 校验报告（JSON）

建议每次抽查：

- 每个题型 2-5 题
- 重点核对：题干、答案、题目解析、难度、知识点、标签

## 常见问题

- **为什么会出现空白列（Unnamed）？**  
  模板尾部有空表头列。脚本已自动裁剪尾部连续空列。

- **为什么答案分隔符不是英文逗号？**  
  默认是中文顿号 `、`。可通过 `--answer-separator ","` 改回英文逗号。

- **这个工具会消耗 token 吗？**  
  不会。当前流程完全本地运行，不调用在线模型接口。
