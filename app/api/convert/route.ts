import { spawn } from "node:child_process";
import { promises as fs } from "node:fs";
import os from "node:os";
import path from "node:path";

export const runtime = "nodejs";

function safeOutputName(raw: string, fallbackStem: string): string {
  const trimmed = (raw || "").trim();
  const base = trimmed || `${fallbackStem}_导出.xlsx`;
  const withExt = base.toLowerCase().endsWith(".xlsx") ? base : `${base}.xlsx`;
  return withExt.replace(/[\\/:*?"<>|]/g, "_");
}

function extractSummary(stdoutText: string): string {
  const m = stdoutText.match(/\{[^{}]*"total"[^{}]*"pass_rate"[^{}]*\}/g);
  return m?.at(-1) ?? "";
}

export async function POST(request: Request): Promise<Response> {
  const formData = await request.formData();
  const wordFile = formData.get("wordFile");
  const templateFile = formData.get("templateFile");
  const answerSeparator = String(formData.get("separator") ?? "、");
  const outputName = String(formData.get("outputName") ?? "");

  if (!(wordFile instanceof File)) {
    return Response.json({ error: "缺少 Word 文件" }, { status: 400 });
  }
  if (!(templateFile instanceof File)) {
    return Response.json({ error: "缺少模板文件" }, { status: 400 });
  }
  if (!["、", ",", "，", ""].includes(answerSeparator)) {
    return Response.json({ error: "答案分隔符不支持" }, { status: 400 });
  }

  const projectRoot = path.join(/* turbopackIgnore: true */ process.cwd());
  const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "word2excel-"));
  const inputPath = path.join(tempRoot, wordFile.name || "input.docx");
  const templatePath = path.join(tempRoot, templateFile.name || "template.xlsx");
  const outputPath = path.join(
    tempRoot,
    safeOutputName(outputName, (wordFile.name || "result").replace(/\.docx$/i, "")),
  );
  const scriptPath = path.join(projectRoot, "word_to_questionbank_excel.py");

  await fs.writeFile(inputPath, Buffer.from(await wordFile.arrayBuffer()));
  await fs.writeFile(templatePath, Buffer.from(await templateFile.arrayBuffer()));

  const args = [
    scriptPath,
    "--input",
    inputPath,
    "--template",
    templatePath,
    "--output",
    outputPath,
    "--module",
    "默认标签",
    "--answer-separator",
    answerSeparator,
  ];

  const proc = spawn("python", args, { cwd: projectRoot });
  let stdout = "";
  let stderr = "";
  proc.stdout.on("data", (chunk) => {
    stdout += String(chunk);
  });
  proc.stderr.on("data", (chunk) => {
    stderr += String(chunk);
  });

  const exitCode: number = await new Promise((resolve) => {
    proc.on("close", resolve);
  });

  if (exitCode !== 0) {
    await fs.rm(tempRoot, { recursive: true, force: true });
    return Response.json(
      {
        error: "导出失败",
        detail: `${stdout}\n${stderr}`.trim(),
      },
      { status: 500 },
    );
  }

  const fileBytes = await fs.readFile(outputPath);
  const summary = extractSummary(stdout);
  await fs.rm(tempRoot, { recursive: true, force: true });

  return new Response(fileBytes, {
    status: 200,
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      "Content-Disposition": `attachment; filename*=UTF-8''${encodeURIComponent(path.basename(outputPath))}`,
      "X-Export-Summary": encodeURIComponent(summary),
      "X-Export-Logs": encodeURIComponent(`${stdout}\n${stderr}`.trim()),
    },
  });
}

