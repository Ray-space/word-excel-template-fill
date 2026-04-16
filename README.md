# Word Excel Template Fill

面向题库导入场景的结构化转换工具：将试题型 `Word(.docx)` 按 `Excel` 导题模板自动解析、分类并导出为 `.xlsx`，同时附带可复用的 Cursor Skill。

这个项目不是泛化的“文档转表格”脚本，而是一个针对教育题库导入、模板对齐、字段结构化抽取而设计的垂直场景工具，重点解决长文本题库在实际导入流程中的稳定性和可核验性问题。

## 产品定位

`Word Excel Template Fill` 解决的是一个高频但高度重复的运营问题：题库资料常常保存在 `Word` 文档中，而目标平台要求按固定的 `Excel` 模板批量导入。人工整理不仅慢，而且容易在题型、答案、解析、知识点等字段上出错。

在这类长文本场景下，单纯依赖大语言模型或 agent 也并不稳定，常见问题包括上下文过长、结构识别漂移、字段遗漏，以及批量处理时出现“数据坍塌”或结果前后不一致，最后仍然需要大量人工复核。

这个项目会读取 Word 题目内容，识别题干、选项、答案、解析、难度、知识点等结构化字段，再按模板表头语义写回 Excel，适合把“导题”从人工整理或不稳定的 AI 提取，变成可重复执行、可抽查、可验证的标准流程。

它同时包含一个 Cursor Skill，适合希望把“Word 转 Excel 导题流程”沉淀为团队标准工作流的场景。

## 核心价值

- 面向题库平台导入模板，而不是通用文档转换
- 本地运行，不依赖在线模型或外部 API
- 识别题型、选项、答案、解析等业务字段，而不只是转格式
- 可作为脚本工具使用，也可作为 Cursor Skill 复用
- 输入输出清晰，适合二次集成、定制和批量处理

## 适合谁用

- 教培机构的题库运营人员
- 需要批量整理导题模板的教研团队
- 做题库迁移、内容清洗、结构化入库的项目成员
- 希望把重复导题流程沉淀为内部工具或 Skill 的团队

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

## 为什么不是普通小脚本

- 关注的是“题库模板对齐”而不是“文件格式转换”
- 内置了题型识别、答案规范化、字段抽取等业务规则
- 既能直接运行，也能通过 Skill 集成到 Cursor 工作流
- 适合在真实导题场景中长期复用，而不是一次性处理单个文件

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
