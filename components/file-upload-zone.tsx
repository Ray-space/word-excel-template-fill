"use client";

import { useCallback, useState } from "react";
import { Upload, FileText, X, Check } from "lucide-react";
import { cn } from "@/lib/utils";

interface FileUploadZoneProps {
  title: string;
  accept: string;
  fileType: string;
  onFileSelect: (file: File | null) => void;
  selectedFile: File | null;
}

export function FileUploadZone({
  title,
  accept,
  fileType,
  onFileSelect,
  selectedFile,
}: FileUploadZoneProps) {
  const [isDragOver, setIsDragOver] = useState(false);

  const handleDragOver = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(true);
  }, []);

  const handleDragLeave = useCallback((e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);
  }, []);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setIsDragOver(false);
      const file = e.dataTransfer.files[0];
      if (file) {
        onFileSelect(file);
      }
    },
    [onFileSelect]
  );

  const handleFileChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const file = e.target.files?.[0];
      if (file) {
        onFileSelect(file);
      }
    },
    [onFileSelect]
  );

  const handleRemoveFile = useCallback(() => {
    onFileSelect(null);
  }, [onFileSelect]);

  return (
    <div className="space-y-3">
      <div className="flex items-center gap-2">
        <div className="flex items-center justify-center w-8 h-8 rounded-lg bg-primary/10">
          <FileText className="w-4 h-4 text-primary" />
        </div>
        <span className="text-sm font-medium text-foreground">{title}</span>
        <span className="text-xs text-muted-foreground">({fileType})</span>
      </div>

      <div
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        className={cn(
          "relative group rounded-xl border-2 border-dashed transition-all duration-300 cursor-pointer",
          "hover:border-primary/50 hover:bg-primary/5",
          isDragOver
            ? "border-primary bg-primary/10 scale-[1.02]"
            : "border-border bg-card/50",
          selectedFile && "border-accent bg-accent/5"
        )}
      >
        <input
          type="file"
          accept={accept}
          onChange={handleFileChange}
          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
        />

        {selectedFile ? (
          <div className="flex items-center justify-between p-4">
            <div className="flex items-center gap-3">
              <div className="flex items-center justify-center w-10 h-10 rounded-lg bg-accent/20">
                <Check className="w-5 h-5 text-accent" />
              </div>
              <div>
                <p className="text-sm font-medium text-foreground truncate max-w-[200px]">
                  {selectedFile.name}
                </p>
                <p className="text-xs text-muted-foreground">
                  {(selectedFile.size / 1024).toFixed(1)} KB
                </p>
              </div>
            </div>
            <button
              onClick={(e) => {
                e.stopPropagation();
                handleRemoveFile();
              }}
              className="relative z-20 p-2 rounded-lg hover:bg-destructive/10 transition-colors"
            >
              <X className="w-4 h-4 text-destructive" />
            </button>
          </div>
        ) : (
          <div className="flex flex-col items-center justify-center py-8 px-4">
            <div
              className={cn(
                "flex items-center justify-center w-14 h-14 rounded-2xl mb-4 transition-all duration-300",
                "bg-gradient-to-br from-primary/20 to-primary/5",
                "group-hover:from-primary/30 group-hover:to-primary/10 group-hover:scale-110",
                isDragOver && "scale-110 from-primary/40 to-primary/20"
              )}
            >
              <Upload
                className={cn(
                  "w-6 h-6 text-primary transition-transform duration-300",
                  "group-hover:-translate-y-0.5",
                  isDragOver && "-translate-y-1"
                )}
              />
            </div>
            <p className="text-sm text-foreground mb-1">将文件拖放到此处</p>
            <p className="text-xs text-muted-foreground">
              或{" "}
              <span className="text-primary hover:underline">点击上传</span>
            </p>
          </div>
        )}
      </div>
    </div>
  );
}
