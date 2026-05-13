"use client";

import { useState } from "react";
import { FileUploadZone } from "@/components/file-upload-zone";
import { Button } from "@/components/ui/button";
import { Input } from "@/components/ui/input";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "@/components/ui/select";
import {
  Download,
  FileOutput,
  Settings2,
  ArrowRight,
  CheckCircle2,
} from "lucide-react";
import { cn } from "@/lib/utils";

export default function HomePage() {
  const [wordFile, setWordFile] = useState<File | null>(null);
  const [templateFile, setTemplateFile] = useState<File | null>(null);
  const [outputFileName, setOutputFileName] = useState("");
  const [separator, setSeparator] = useState("、");
  const [isExporting, setIsExporting] = useState(false);
  const [exportComplete, setExportComplete] = useState(false);
  const [exportedFilePath, setExportedFilePath] = useState("");
  const [downloadBlob, setDownloadBlob] = useState<Blob | null>(null);
  const [errorText, setErrorText] = useState("");
  const [summaryText, setSummaryText] = useState("");

  const canExport = wordFile && templateFile;

  const handleExport = async () => {
    if (!canExport) return;

    setErrorText("");
    setSummaryText("");
    setExportComplete(false);
    setIsExporting(true);
    try {
      const formData = new FormData();
      formData.append("wordFile", wordFile);
      formData.append("templateFile", templateFile);
      formData.append("outputName", outputFileName);
      const normalizedSeparator = separator === "none" ? "" : separator;
      formData.append("separator", normalizedSeparator);

      const resp = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!resp.ok) {
        const errData = (await resp.json()) as { error?: string; detail?: string };
        throw new Error(errData.detail || errData.error || "导出失败");
      }

      const blob = await resp.blob();
      const nameFromHeader =
        resp.headers
          .get("Content-Disposition")
          ?.match(/filename\*=UTF-8''(.+)$/)?.[1] || "";
      const decodedFileName = nameFromHeader
        ? decodeURIComponent(nameFromHeader)
        : outputFileName || `${wordFile.name.replace(".docx", "")}_导出.xlsx`;

      const summary = resp.headers.get("X-Export-Summary");
      if (summary) setSummaryText(decodeURIComponent(summary));

      setDownloadBlob(blob);
      setExportedFilePath(decodedFileName);
      setExportComplete(true);
    } catch (err) {
      setErrorText(err instanceof Error ? err.message : "导出失败");
    } finally {
      setIsExporting(false);
    }
  };

  const handleDownload = () => {
    if (!downloadBlob) return;
    const url = URL.createObjectURL(downloadBlob);
    const link = document.createElement("a");
    link.href = url;
    link.download = exportedFilePath || "导出结果.xlsx";
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  const handleReset = () => {
    setWordFile(null);
    setTemplateFile(null);
    setOutputFileName("");
    setSeparator("、");
    setExportComplete(false);
    setExportedFilePath("");
    setDownloadBlob(null);
    setErrorText("");
    setSummaryText("");
  };

  return (
    <main className="min-h-screen bg-background">
      {/* 背景装饰 */}
      <div className="fixed inset-0 overflow-hidden pointer-events-none">
        <div className="absolute -top-40 -right-40 w-80 h-80 bg-primary/10 rounded-full blur-3xl" />
        <div className="absolute -bottom-40 -left-40 w-80 h-80 bg-accent/10 rounded-full blur-3xl" />
        <div className="absolute top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[600px] h-[600px] bg-primary/5 rounded-full blur-3xl" />
      </div>

      <div className="relative z-10 container mx-auto px-4 py-8 lg:py-12 max-w-3xl">
        {/* 头部 */}
        <header className="text-center mb-10">
          {/* 品牌 Logo */}
          <div className="inline-flex items-center gap-4 mb-6">
            {/* 精致图标 */}
            <div className="relative group">
              <div className="absolute inset-0 bg-gradient-to-br from-primary/60 to-accent/60 rounded-2xl blur-lg opacity-50 group-hover:opacity-70 transition-opacity" />
              <div className="relative w-12 h-12 rounded-2xl bg-gradient-to-br from-card to-secondary border border-border/50 flex items-center justify-center overflow-hidden">
                {/* 内部装饰线条 */}
                <div className="absolute inset-0">
                  <div className="absolute top-2 left-2 w-3 h-[2px] bg-gradient-to-r from-primary to-transparent rounded-full" />
                  <div className="absolute top-2 left-2 w-[2px] h-3 bg-gradient-to-b from-primary to-transparent rounded-full" />
                  <div className="absolute bottom-2 right-2 w-3 h-[2px] bg-gradient-to-l from-accent to-transparent rounded-full" />
                  <div className="absolute bottom-2 right-2 w-[2px] h-3 bg-gradient-to-t from-accent to-transparent rounded-full" />
                </div>
                {/* 中心箭头图标 */}
                <svg className="w-5 h-5 text-foreground relative z-10" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="2" strokeLinecap="round" strokeLinejoin="round">
                  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" className="stroke-primary" />
                  <polyline points="14 2 14 8 20 8" className="stroke-primary" />
                  <path d="M12 18v-6" className="stroke-accent" />
                  <path d="M9 15l3 3 3-3" className="stroke-accent" />
                </svg>
              </div>
            </div>
            
            {/* 品牌名称 - 艺术字体 */}
            <div className="text-left">
              <h2 className="text-xl font-bold tracking-tight text-foreground" style={{ fontFamily: 'var(--font-brand)' }}>
                <span className="bg-gradient-to-r from-foreground via-foreground to-muted-foreground bg-clip-text">
                  run
                </span>
                <span className="text-primary">_</span>
                <span className="bg-gradient-to-r from-primary to-accent bg-clip-text text-transparent">
                  word
                </span>
                <span className="text-muted-foreground">_to_</span>
                <span className="bg-gradient-to-r from-accent to-primary bg-clip-text text-transparent">
                  excel
                </span>
              </h2>
              <p className="text-[10px] text-muted-foreground/70 tracking-widest uppercase mt-0.5">
                Document Converter
              </p>
            </div>
          </div>

          <h1 className="text-xl lg:text-2xl font-medium text-foreground mb-2 text-balance">
            快速将 Word 文档转换为 Excel 格式
          </h1>
          <p className="text-muted-foreground text-sm max-w-md mx-auto">
            上传文件并配置参数，一键导出结果
          </p>
        </header>

        {/* 主要内容 */}
        <div className="space-y-6">
          {/* 文件上传区域 */}
          <div className="grid gap-6 lg:grid-cols-2">
            <FileUploadZone
              title="上传 Word 文件"
              accept=".docx"
              fileType=".docx"
              onFileSelect={setWordFile}
              selectedFile={wordFile}
            />
            <FileUploadZone
              title="上传模板文件"
              accept=".xlsx"
              fileType=".xlsx"
              onFileSelect={setTemplateFile}
              selectedFile={templateFile}
            />
          </div>

          {/* 设置区域 */}
          <div className="rounded-xl border border-border bg-card/50 backdrop-blur-sm p-5">
            <div className="flex items-center gap-2 mb-4">
              <Settings2 className="w-4 h-4 text-primary" />
              <span className="text-sm font-medium text-foreground">
                导出设置
              </span>
            </div>

            <div className="grid gap-4 lg:grid-cols-2">
              <div className="space-y-2">
                <label className="text-xs text-muted-foreground">
                  输出文件名（可选）
                </label>
                <Input
                  placeholder="例如：选判断题10道_按试题模版导出.xlsx"
                  value={outputFileName}
                  onChange={(e) => setOutputFileName(e.target.value)}
                  className="bg-input border-border focus:border-primary"
                />
              </div>

              <div className="space-y-2">
                <label className="text-xs text-muted-foreground">
                  多选答案分隔符
                </label>
                <Select value={separator} onValueChange={setSeparator}>
                  <SelectTrigger className="bg-input border-border">
                    <SelectValue />
                  </SelectTrigger>
                  <SelectContent>
                    <SelectItem value="、">顿号 (、)</SelectItem>
                    <SelectItem value=",">英文逗号 (,)</SelectItem>
                    <SelectItem value="，">中文逗号 (，)</SelectItem>
                    <SelectItem value="none">无分隔符</SelectItem>
                  </SelectContent>
                </Select>
              </div>
            </div>
          </div>
          {errorText && (
            <div className="rounded-xl border border-destructive/40 bg-destructive/10 p-4 text-sm text-destructive">
              {errorText}
            </div>
          )}

          {/* 导出按钮 */}
          <Button
            onClick={handleExport}
            disabled={!canExport || isExporting}
            className={cn(
              "w-full h-12 text-base font-medium transition-all duration-300",
              "bg-gradient-to-r from-primary to-primary/80",
              "hover:from-primary/90 hover:to-primary/70 hover:shadow-lg hover:shadow-primary/25",
              "disabled:opacity-50 disabled:cursor-not-allowed disabled:hover:shadow-none"
            )}
          >
            {isExporting ? (
              <div className="flex items-center gap-2">
                <div className="w-4 h-4 border-2 border-primary-foreground/30 border-t-primary-foreground rounded-full animate-spin" />
                <span>正在导出...</span>
              </div>
            ) : (
              <div className="flex items-center gap-2">
                <FileOutput className="w-5 h-5" />
                <span>开始导出</span>
                <ArrowRight className="w-4 h-4" />
              </div>
            )}
          </Button>

          {/* 导出结果 */}
          {exportComplete && (
            <div className="rounded-xl border border-accent/30 bg-accent/5 backdrop-blur-sm p-5 animate-in fade-in slide-in-from-bottom-4 duration-500">
              <div className="flex items-center gap-2 mb-4">
                <CheckCircle2 className="w-5 h-5 text-accent" />
                <span className="text-sm font-medium text-foreground">
                  导出完成
                </span>
              </div>

              <div className="space-y-4">
                <div className="flex items-center gap-3 p-3 rounded-lg bg-card/80 border border-border">
                  <div className="flex items-center justify-center w-10 h-10 rounded-lg bg-accent/20">
                    <FileOutput className="w-5 h-5 text-accent" />
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-medium text-foreground truncate">
                      {exportedFilePath}
                    </p>
                    <p className="text-xs text-muted-foreground">
                      文件已准备就绪
                    </p>
                    {summaryText && (
                      <p className="text-xs text-muted-foreground mt-1 truncate">
                        校验摘要：{summaryText}
                      </p>
                    )}
                  </div>
                </div>

                <div className="flex gap-3">
                  <Button
                    onClick={handleDownload}
                    className="flex-1 bg-accent hover:bg-accent/90 text-accent-foreground"
                  >
                    <Download className="w-4 h-4 mr-2" />
                    下载导出文件
                  </Button>
                  <Button
                    onClick={handleReset}
                    variant="outline"
                    className="border-border hover:bg-secondary"
                  >
                    重新开始
                  </Button>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* 底部 */}
        <footer className="mt-12 pt-6 border-t border-border/50">
          <div className="flex items-center justify-center gap-3 text-xs text-muted-foreground">
            <span className="font-semibold tracking-tight" style={{ fontFamily: 'var(--font-brand)' }}>
              run<span className="text-primary">_</span>word<span className="text-muted-foreground/50">_to_</span>excel
            </span>
            <span className="w-1 h-1 rounded-full bg-border" />
            <span>支持 .docx 与 .xlsx</span>
            <span className="w-1 h-1 rounded-full bg-border" />
            <span>本地处理</span>
          </div>
        </footer>
      </div>
    </main>
  );
}
