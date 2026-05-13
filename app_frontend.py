import json
import re
import subprocess
import sys
from pathlib import Path
from typing import Optional, Tuple

import gradio as gr


ROOT = Path(__file__).resolve().parent
ENTRY_SCRIPT = ROOT / "word_to_questionbank_excel.py"
DEFAULT_OUTPUT_DIR = ROOT / "frontend_outputs"


def _extract_summary(stdout_text: str) -> Optional[dict]:
    matches = re.findall(r"\{[^{}]*\"total\"[^{}]*\"pass_rate\"[^{}]*\}", stdout_text)
    if not matches:
        return None
    try:
        return json.loads(matches[-1])
    except json.JSONDecodeError:
        return None


def _safe_module_name(text: str) -> str:
    cleaned = re.sub(r"[\\/:*?\"<>|]+", "_", (text or "").strip())
    return cleaned or "未命名模块"


def run_convert(
    input_docx_file: Optional[str],
    template_xlsx_file: Optional[str],
    output_name: str,
    answer_separator: str,
) -> Tuple[str, Optional[str], Optional[str], gr.update]:
    if not input_docx_file:
        return (
            "请先上传 Word 文件（.docx）。",
            None,
            None,
            gr.update(value=None, interactive=False),
        )
    if not template_xlsx_file:
        return (
            "请先上传模板文件（.xlsx）。",
            None,
            None,
            gr.update(value=None, interactive=False),
        )
    if not ENTRY_SCRIPT.exists():
        return (
            f"入口脚本不存在: {ENTRY_SCRIPT}",
            None,
            None,
            gr.update(value=None, interactive=False),
        )
    module_name = "默认标签"

    input_path = Path(input_docx_file)
    template_path = Path(template_xlsx_file)

    if not input_path.exists():
        return (
            f"Word 文件不存在: {input_path}",
            None,
            None,
            gr.update(value=None, interactive=False),
        )
    if not template_path.exists():
        return (
            f"模板文件不存在: {template_path}",
            None,
            None,
            gr.update(value=None, interactive=False),
        )

    DEFAULT_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    if output_name.strip():
        output_filename = output_name.strip()
        if not output_filename.lower().endswith(".xlsx"):
            output_filename += ".xlsx"
    else:
        output_filename = f"{input_path.stem}_{_safe_module_name(module_name)}_导出.xlsx"
    output_path = DEFAULT_OUTPUT_DIR / output_filename

    cmd = [
        sys.executable,
        str(ENTRY_SCRIPT),
        "--input",
        str(input_path),
        "--template",
        str(template_path),
        "--output",
        str(output_path),
        "--module",
        module_name,
        "--answer-separator",
        answer_separator,
    ]

    proc = subprocess.run(cmd, capture_output=True, text=True)
    stdout = proc.stdout or ""
    stderr = proc.stderr or ""
    summary = _extract_summary(stdout)

    report_lines = []
    report_lines.append("## 运行结果")
    report_lines.append(f"- 退出码: `{proc.returncode}`")
    report_lines.append(f"- 输出文件: `{output_path}`")
    report_lines.append("- 下载方式: 在页面底部 `导出结果下载` 组件点击下载，可自行选择保存目录。")
    if summary:
        report_lines.append(
            "- 校验摘要: "
            f"`total={summary.get('total')} | critical={summary.get('critical')} | "
            f"warnings={summary.get('warnings')} | pass_rate={summary.get('pass_rate')}`"
        )
    report_lines.append("")
    report_lines.append("## 控制台输出")
    report_lines.append("```text")
    report_lines.append((stdout + "\n" + stderr).strip() or "(无输出)")
    report_lines.append("```")

    if proc.returncode != 0 or not output_path.exists():
        return (
            "\n".join(report_lines),
            None,
            None,
            gr.update(value=None, interactive=False),
        )
    return (
        "\n".join(report_lines),
        str(output_path),
        str(output_path),
        gr.update(value=str(output_path), interactive=True),
    )


CSS = """
.gradio-container {
  background: radial-gradient(1200px 400px at 20% -10%, #182d55 0%, #0a0f1f 60%, #090d1a 100%);
}
#panel {
  border: 1px solid rgba(76, 119, 255, 0.45);
  border-radius: 16px;
  background: rgba(11, 18, 38, 0.72);
  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.35);
}
#title {
  color: #d6e8ff;
  letter-spacing: 0.5px;
}
#hint {
  color: #9cb8e8;
}
"""


with gr.Blocks(title="Word 导题前端") as demo:
    gr.Markdown(
        "## Word -> Excel 导题前端\n"
        "上传路径后点击导出，按模板表头语义生成新文件。",
        elem_id="title",
    )
    gr.Markdown(
        "上传 Word 与模板后点击导出，结果会直接在页面提供下载。",
        elem_id="hint",
    )
    with gr.Group(elem_id="panel"):
        input_docx = gr.File(
            label="上传 Word 文件 (.docx)",
            file_types=[".docx"],
            type="filepath",
        )
        template_xlsx = gr.File(
            label="上传模板文件 (.xlsx)",
            file_types=[".xlsx"],
            type="filepath",
        )
        output_name = gr.Textbox(
            label="输出文件名（可选）",
            placeholder="例如：选判断题10道_按试题模版导出.xlsx",
        )
        answer_separator = gr.Dropdown(
            label="多选答案分隔符",
            choices=["、", ",", "，", ""],
            value="、",
        )
        run_btn = gr.Button("开始导出", variant="primary")
    log_md = gr.Markdown(label="运行日志")
    output_file = gr.File(label="导出结果下载")
    output_path_text = gr.Textbox(label="导出文件路径", interactive=False)
    download_btn = gr.DownloadButton("下载导出文件", interactive=False)

    run_btn.click(
        fn=run_convert,
        inputs=[input_docx, template_xlsx, output_name, answer_separator],
        outputs=[log_md, output_file, output_path_text, download_btn],
    )


if __name__ == "__main__":
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=True,
        theme=gr.themes.Soft(primary_hue="blue", neutral_hue="slate"),
        css=CSS,
    )
