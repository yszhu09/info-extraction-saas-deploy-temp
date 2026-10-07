"use client";

import Link from "next/link";
import { useEffect, useRef, useState } from "react";
import { JeStatusBadge } from "@/components/JeStatusBadge";
import type { ClipboardEvent, DragEvent } from "react";
import ExcelJS from "exceljs";
import {
  isSpreadsheetFile,
  parseSpreadsheetFile,
  SPREADSHEET_ACCEPT,
} from "@/lib/spreadsheets";
import type {
  ExtractRequestPayload,
  ExtractResponsePayload,
  SourceFilePayload,
  SourceImage,
  SourceTable,
} from "@/lib/types";

type BusyState = "source" | "template" | "extract" | null;

const MAX_MODEL_IMAGES = 8;
const MAX_IMAGE_DIMENSION = 1400;
const IMAGE_JPEG_QUALITY = 0.72;
const MIN_IMAGE_DIMENSION = 16;

function fileToDataUrl(file: File) {
  return new Promise<string>((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result));
    reader.onerror = () => reject(reader.error ?? new Error("文件读取失败"));
    reader.readAsDataURL(file);
  });
}

function loadImage(dataUrl: string) {
  return new Promise<HTMLImageElement>((resolve, reject) => {
    const image = new Image();
    image.onload = () => resolve(image);
    image.onerror = () => reject(new Error("图片读取失败，请换一张图片。"));
    image.src = dataUrl;
  });
}

async function imageFileToSourceImage(
  file: File,
  fallbackName = "clipboard-image.png",
): Promise<SourceImage> {
  const name = file.name || fallbackName;
  const image = await loadImage(await fileToDataUrl(file));
  const width = image.naturalWidth || image.width;
  const height = image.naturalHeight || image.height;

  if (width < MIN_IMAGE_DIMENSION || height < MIN_IMAGE_DIMENSION) {
    throw new Error(`图片尺寸过小，已跳过：${name}`);
  }

  const scale = Math.min(1, MAX_IMAGE_DIMENSION / Math.max(width, height));
  const targetWidth = Math.max(1, Math.round(width * scale));
  const targetHeight = Math.max(1, Math.round(height * scale));
  const canvas = document.createElement("canvas");
  const context = canvas.getContext("2d", { alpha: false });
  if (!context) throw new Error("浏览器无法压缩图片。");

  canvas.width = targetWidth;
  canvas.height = targetHeight;
  context.fillStyle = "#ffffff";
  context.fillRect(0, 0, targetWidth, targetHeight);
  context.drawImage(image, 0, 0, targetWidth, targetHeight);
  const dataUrl = canvas.toDataURL("image/jpeg", IMAGE_JPEG_QUALITY);
  canvas.width = 1;
  canvas.height = 1;

  return { name, mimeType: "image/jpeg", dataUrl };
}

function countPayloadImages(payload: ExtractRequestPayload) {
  return (
    (payload.template.screenshotTemplate ? 1 : 0) +
    payload.files.reduce((count, file) => count + (file.images?.length ?? 0), 0)
  );
}

async function loadPdfJs() {
  const pdfjs = await import("pdfjs-dist");
  pdfjs.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";
  return pdfjs;
}

async function parsePdfFile(file: File): Promise<SourceFilePayload> {
  const pdfjs = await loadPdfJs();
  const bytes = await file.arrayBuffer();
  const pdf = await pdfjs.getDocument({ data: bytes.slice(0) }).promise;
  if (pdf.numPages > 20) {
    throw new Error(
      `PDF 共 ${pdf.numPages} 页，MVP 最多处理 20 页；请先拆分或缩小页码范围。`,
    );
  }

  const pageTexts: string[] = [];
  const images: SourceImage[] = [];
  for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
    const page = await pdf.getPage(pageNumber);
    const textContent = await page.getTextContent();
    const text = textContent.items
      .map((item) => String((item as { str?: string }).str ?? ""))
      .join(" ")
      .replace(/\s+/g, " ")
      .trim();
    pageTexts.push(
      `Page ${pageNumber}: ${text || "（未提取到文字，已附页面图片）"}`,
    );

    const viewport = page.getViewport({ scale: 1.35 });
    const canvas = document.createElement("canvas");
    const context = canvas.getContext("2d", { alpha: false });
    if (!context) throw new Error("浏览器无法渲染 PDF 页面图片。");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvas, canvasContext: context, viewport }).promise;
    images.push({
      name: `${file.name} page ${pageNumber}`,
      mimeType: "image/jpeg",
      dataUrl: canvas.toDataURL("image/jpeg", 0.72),
    });
    canvas.width = 1;
    canvas.height = 1;
    page.cleanup();
  }

  return {
    kind: "pdf",
    name: file.name,
    pageCount: pdf.numPages,
    text: pageTexts.join("\n"),
    images,
  };
}

async function processSourceFile(file: File): Promise<SourceFilePayload> {
  const lowerName = file.name.toLowerCase();
  if (file.type.startsWith("image/")) {
    const image = await imageFileToSourceImage(file);
    return {
      kind: "image",
      name: image.name,
      text: `图片/截图文件：${image.name}（已压缩为模型可读图片）`,
      images: [image],
    };
  }
  if (file.type === "application/pdf" || lowerName.endsWith(".pdf")) {
    return parsePdfFile(file);
  }
  if (isSpreadsheetFile(file)) {
    const tables = await parseSpreadsheetFile(file);
    return {
      kind: "excel",
      name: file.name,
      text: tables
        .map(
          (table) =>
            `${table.sheetName}: ${table.columns.join(" | ")}\n${table.rows.map((row) => row.join(" | ")).join("\n")}`,
        )
        .join("\n\n"),
      tables,
    };
  }
  throw new Error(`暂不支持该文件类型：${file.name}`);
}

function toTsv(columns: string[], rows: string[][]) {
  const escapeCell = (value: string) =>
    value.replace(/\t/g, " ").replace(/\r?\n/g, " ");
  return [columns, ...rows]
    .map((row) => row.map(escapeCell).join("\t"))
    .join("\n");
}

function bufferToArrayBuffer(buffer: ExcelJS.Buffer) {
  if (buffer instanceof ArrayBuffer) return buffer;
  const view = buffer as Uint8Array;
  return view.buffer.slice(
    view.byteOffset,
    view.byteOffset + view.byteLength,
  ) as ArrayBuffer;
}

export function ExtractorApp() {
  const [sourceText, setSourceText] = useState("");
  const [sourceFiles, setSourceFiles] = useState<SourceFilePayload[]>([]);
  const [instruction, setInstruction] = useState("");
  const [excelTemplate, setExcelTemplate] = useState<SourceTable | undefined>();
  const [screenshotTemplate, setScreenshotTemplate] = useState<
    SourceImage | undefined
  >();
  const [result, setResult] = useState<ExtractResponsePayload | null>(null);
  const [busy, setBusy] = useState<BusyState>(null);
  const [status, setStatus] = useState("请先添加资料和提取要求。");
  const [error, setError] = useState("");
  const sourceInputRef = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const sourceTextAreaRef = useRef<HTMLTextAreaElement>(null);
  const templateTextAreaRef = useRef<HTMLTextAreaElement>(null);

  const hasSource = sourceFiles.length > 0 || sourceText.trim().length > 0;
  const hasTemplate = Boolean(
    instruction.trim() || excelTemplate || screenshotTemplate,
  );

  async function handleSourceFiles(files: FileList | File[] | null) {
    const nextFiles = Array.from(files ?? []);
    if (!nextFiles.length) return;
    setBusy("source");
    setError("");
    try {
      const processed: SourceFilePayload[] = [];
      const skipped: string[] = [];
      for (const file of nextFiles) {
        try {
          processed.push(await processSourceFile(file));
        } catch (fileError) {
          skipped.push(
            fileError instanceof Error
              ? fileError.message
              : `${file.name} 处理失败`,
          );
        }
      }
      if (processed.length) {
        setSourceFiles((current) => [...current, ...processed]);
        setStatus(
          `已添加 ${processed.length} 个资料文件；图片会自动压缩，提取时最多发送前 ${MAX_MODEL_IMAGES} 张图片。`,
        );
      }
      if (skipped.length) {
        setError(`部分文件未添加：${skipped.slice(0, 3).join("；")}`);
      }
      if (!processed.length && !skipped.length)
        setStatus("未检测到可处理文件。");
    } catch (fileError) {
      setError(fileError instanceof Error ? fileError.message : "文件处理失败");
    } finally {
      setBusy(null);
    }
  }

  async function handleExcelTemplate(file: File | null) {
    if (!file) return;
    setBusy("template");
    setError("");
    try {
      const [table] = await parseSpreadsheetFile(file);
      setExcelTemplate(table);
      setStatus(`已读取 Excel / CSV 模板：${table.columns.join(" / ")}`);
    } catch (templateError) {
      setError(
        templateError instanceof Error
          ? templateError.message
          : "Excel / CSV 模板解析失败",
      );
    } finally {
      setBusy(null);
    }
  }

  async function handleScreenshotTemplate(file: File | null) {
    if (!file) return;
    setBusy("template");
    setError("");
    try {
      const image = await imageFileToSourceImage(
        file,
        "clipboard-template.png",
      );
      setScreenshotTemplate(image);
      setStatus(`已添加模板图片：${image.name}；已压缩为模型可读图片。`);
    } catch (templateError) {
      setError(
        templateError instanceof Error
          ? templateError.message
          : "模板图片读取失败",
      );
    } finally {
      setBusy(null);
    }
  }

  async function handleTemplateFiles(files: FileList | File[] | null) {
    const nextFiles = Array.from(files ?? []);
    if (!nextFiles.length) return;
    const excelFile = nextFiles.find((file) =>
      isSpreadsheetFile(file),
    );
    const imageFile = nextFiles.find((file) => file.type.startsWith("image/"));
    if (!excelFile && !imageFile) {
      setError("模板区支持 Excel / CSV 模板或图片/截图模板。");
      return;
    }
    if (excelFile) await handleExcelTemplate(excelFile);
    if (imageFile) await handleScreenshotTemplate(imageFile);
  }

  function clipboardImagesFromData(clipboardData: DataTransfer | null) {
    if (!clipboardData) return [];
    const files = Array.from(clipboardData.files).filter((file) =>
      file.type.startsWith("image/"),
    );
    if (files.length) return files;
    return Array.from(clipboardData.items)
      .filter((item) => item.type.startsWith("image/"))
      .map((item) => item.getAsFile())
      .filter((file): file is File => Boolean(file));
  }

  function clipboardImages(event: ClipboardEvent<HTMLElement>) {
    return clipboardImagesFromData(event.clipboardData);
  }

  async function handleSourcePaste(event: ClipboardEvent<HTMLElement>) {
    const images = clipboardImages(event);
    if (!images.length) return;
    event.preventDefault();
    await handleSourceFiles(images);
  }

  async function handleTemplatePaste(event: ClipboardEvent<HTMLElement>) {
    const images = clipboardImages(event);
    if (!images.length) return;
    event.preventDefault();
    await handleScreenshotTemplate(images[0]);
  }

  function stopDrag(event: DragEvent<HTMLElement>) {
    event.preventDefault();
    event.stopPropagation();
  }

  async function handleSourceDrop(event: DragEvent<HTMLElement>) {
    stopDrag(event);
    await handleSourceFiles(event.dataTransfer.files);
  }

  async function handleTemplateDrop(event: DragEvent<HTMLElement>) {
    stopDrag(event);
    await handleTemplateFiles(event.dataTransfer.files);
  }

  useEffect(() => {
    const sourceElement = sourceTextAreaRef.current;
    const templateElement = templateTextAreaRef.current;

    const handleSourceTextAreaPaste = (event: globalThis.ClipboardEvent) => {
      const images = clipboardImagesFromData(event.clipboardData);
      if (!images.length) return;
      event.preventDefault();
      event.stopPropagation();
      void handleSourceFiles(images);
    };
    const handleTemplateTextAreaPaste = (event: globalThis.ClipboardEvent) => {
      const images = clipboardImagesFromData(event.clipboardData);
      if (!images.length) return;
      event.preventDefault();
      event.stopPropagation();
      void handleScreenshotTemplate(images[0]);
    };

    sourceElement?.addEventListener("paste", handleSourceTextAreaPaste);
    templateElement?.addEventListener("paste", handleTemplateTextAreaPaste);

    return () => {
      sourceElement?.removeEventListener("paste", handleSourceTextAreaPaste);
      templateElement?.removeEventListener(
        "paste",
        handleTemplateTextAreaPaste,
      );
    };
  });

  async function handleExtract() {
    if (!hasSource) {
      setError("请先输入文字资料或上传资料文件。");
      return;
    }
    setBusy("extract");
    setError("");
    setResult(null);
    try {
      const payload: ExtractRequestPayload = {
        sourceText,
        files: sourceFiles,
        template: {
          instruction,
          excelTemplate,
          screenshotTemplate,
          outputMode: "both",
        },
      };
      const imageCount = countPayloadImages(payload);
      const imageStatus = imageCount
        ? imageCount > MAX_MODEL_IMAGES
          ? `图片已压缩；本次优先发送前 ${MAX_MODEL_IMAGES} 张，超出的仅使用文字摘要。`
          : "图片已压缩后发送给模型。"
        : "";
      setStatus(imageStatus || "正在提取并生成 Excel 表格…");
      const response = await fetch("/api/extract", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const body = await response.json();
      if (!response.ok) throw new Error(body.error || "提取失败");
      setResult(body as ExtractResponsePayload);
      setStatus(
        imageStatus
          ? `提取完成，已生成 Excel 表格。${imageStatus}`
          : "提取完成，已生成 Excel 表格。",
      );
    } catch (extractError) {
      setError(
        extractError instanceof Error ? extractError.message : "提取失败",
      );
    } finally {
      setBusy(null);
    }
  }

  async function copyTsv() {
    if (!result) return;
    await navigator.clipboard.writeText(toTsv(result.columns, result.rows));
    setStatus("表格已复制到剪贴板。");
  }

  async function downloadXlsx() {
    if (!result) return;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Extracted Data");
    sheet.addRow(result.columns);
    result.rows.forEach((row) => sheet.addRow(row));
    sheet.columns.forEach((column) => {
      column.width = Math.max(
        12,
        ...((column.values ?? []).map((value) => String(value ?? "").length) as
          | number[]
          | []),
      );
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([bufferToArrayBuffer(buffer)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "extracted-result.xlsx";
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  async function logout() {
    await fetch("/api/logout", { method: "POST" });
    window.location.href = "/login";
  }

  const sourceKindLabel: Record<SourceFilePayload["kind"], string> = {
    pdf: "PDF",
    image: "图片",
    excel: "Excel",
  };

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100" data-infowb-page="extract">
      <div className="mx-auto flex max-w-7xl flex-col gap-5 px-4 py-5 lg:h-screen lg:px-6">
        <header className="flex flex-col gap-3 rounded-3xl border border-white/10 bg-white/8 px-5 py-4 shadow-2xl shadow-black/20 md:flex-row md:items-center md:justify-between">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.28em] text-cyan-300">
              Info Extractor Workbench
            </p>
            <h1 className="mt-1 text-2xl font-bold tracking-tight md:text-3xl">
              信息提取工作台
            </h1>
          </div>
          <div className="flex flex-wrap items-center gap-3 text-sm text-slate-300">
            <nav className="flex rounded-full border border-white/15 bg-slate-900/60 p-1">
              <Link
                prefetch={false}
                href="/"
                className="rounded-full bg-cyan-400 px-3 py-1.5 font-semibold text-slate-950"
              >
                信息提取
              </Link>
              <Link
                prefetch={false}
                href="/compare"
                className="rounded-full px-3 py-1.5 text-slate-300 hover:text-cyan-200"
              >
                信息核对
              </Link>
            </nav>
            <JeStatusBadge page="extract" />
            <span className="rounded-full border border-cyan-300/30 bg-cyan-300/10 px-3 py-1 text-cyan-100">
              上传资料 → 设置要求 → 导出 Excel
            </span>
            <button
              onClick={logout}
              className="rounded-full border border-white/20 px-4 py-2 text-slate-200 hover:border-cyan-300 hover:text-cyan-200"
            >
              退出
            </button>
          </div>
        </header>

        <div className="grid flex-1 gap-5 lg:min-h-0 lg:grid-cols-[430px_minmax(0,1fr)]">
          <section className="flex min-h-0 flex-col gap-4 overflow-auto rounded-3xl border border-white/10 bg-slate-900/70 p-4 shadow-2xl shadow-black/20">
            <div className="rounded-3xl bg-white p-4 text-slate-950 shadow-xl">
              <div className="flex items-center justify-between gap-3">
                <div>
                  <p className="text-xs font-semibold uppercase tracking-[0.22em] text-cyan-700">
                    Source
                  </p>
                  <h2 className="text-xl font-bold">资料</h2>
                </div>
                {sourceFiles.length ? (
                  <button
                    className="rounded-full bg-slate-100 px-3 py-1 text-xs text-slate-600 hover:bg-slate-200"
                    onClick={() => setSourceFiles([])}
                  >
                    清空文件
                  </button>
                ) : null}
              </div>

              <div
                data-testid="extract-source-dropzone"
                data-infowb-dropzone="source"
                onDragEnter={stopDrag}
                onDragOver={stopDrag}
                onDrop={handleSourceDrop}
                onPaste={handleSourcePaste}
                className="mt-4 rounded-3xl border border-slate-200 bg-slate-50 p-3 transition focus-within:border-cyan-500 focus-within:bg-white"
              >
                <div className="relative">
                  <textarea
                    data-testid="extract-source-textarea"
                    data-infowb-textarea="source"
                    ref={sourceTextAreaRef}
                    className="min-h-44 w-full resize-none rounded-2xl bg-transparent p-3 pb-16 text-sm outline-none placeholder:text-slate-400"
                    value={sourceText}
                    onChange={(event) => setSourceText(event.target.value)}
                    placeholder="输入或粘贴资料，也可以粘贴截图、拖入文件，或点回形针上传 PDF / 图片 / Excel / CSV"
                  />
                  <button
                    type="button"
                    aria-label="上传资料文件"
                    onClick={() => sourceInputRef.current?.click()}
                    className="absolute bottom-2 left-2 rounded-full border border-slate-300 bg-white px-3 py-2 text-lg leading-none text-slate-700 shadow-sm transition hover:border-cyan-400 hover:text-cyan-700"
                  >
                    📎
                  </button>
                  <span className="pointer-events-none absolute bottom-4 left-14 text-xs text-slate-400">
                    附件
                  </span>
                </div>
              </div>
              <input
                data-testid="extract-source-file-input"
                data-infowb-file-input="source"
                ref={sourceInputRef}
                className="hidden"
                type="file"
                multiple
                accept={`application/pdf,image/*,${SPREADSHEET_ACCEPT}`}
                onChange={(event) => {
                  void handleSourceFiles(event.target.files);
                  event.currentTarget.value = "";
                }}
              />

              <div className="mt-4 space-y-2">
                {sourceFiles.length ? (
                  sourceFiles.map((file, index) => (
                    <div
                      className="flex items-center justify-between gap-3 rounded-2xl bg-slate-100 px-3 py-2 text-sm text-slate-700"
                      key={`${file.name}-${index}`}
                    >
                      <span className="min-w-0 truncate">
                        {file.name}
                        <span className="ml-2 text-xs text-slate-500">
                          {sourceKindLabel[file.kind]}
                          {file.pageCount ? ` · ${file.pageCount} 页` : ""}
                        </span>
                      </span>
                      <button
                        className="shrink-0 text-xs text-slate-500 hover:text-red-600"
                        onClick={() =>
                          setSourceFiles((current) =>
                            current.filter(
                              (_, itemIndex) => itemIndex !== index,
                            ),
                          )
                        }
                      >
                        移除
                      </button>
                    </div>
                  ))
                ) : (
                  <p className="rounded-2xl bg-slate-100 px-3 py-2 text-sm text-slate-500">
                    还没有上传文件。
                  </p>
                )}
              </div>
            </div>

            <div className="rounded-3xl bg-white p-4 text-slate-950 shadow-xl">
              <div>
                <p className="text-xs font-semibold uppercase tracking-[0.22em] text-cyan-700">
                  Requirement
                </p>
                <h2 className="text-xl font-bold">要求与模板</h2>
              </div>

              <div
                data-testid="extract-template-dropzone"
                data-infowb-dropzone="template"
                onDragEnter={stopDrag}
                onDragOver={stopDrag}
                onDrop={handleTemplateDrop}
                onPaste={handleTemplatePaste}
                className="mt-4 rounded-3xl border border-slate-200 bg-slate-50 p-3 transition focus-within:border-cyan-500 focus-within:bg-white"
              >
                <div className="relative">
                  <textarea
                    data-testid="extract-template-textarea"
                    data-infowb-textarea="template"
                    ref={templateTextAreaRef}
                    className="min-h-40 w-full resize-none rounded-2xl bg-transparent p-3 pb-16 text-sm outline-none placeholder:text-slate-400"
                    value={instruction}
                    onChange={(event) => setInstruction(event.target.value)}
                    placeholder="写清楚要提取哪些字段；也可以粘贴模板截图，或点回形针上传 Excel / CSV 模板"
                  />
                  <button
                    type="button"
                    aria-label="上传要求模板"
                    onClick={() => templateInputRef.current?.click()}
                    className="absolute bottom-2 left-2 rounded-full border border-slate-300 bg-white px-3 py-2 text-lg leading-none text-slate-700 shadow-sm transition hover:border-cyan-400 hover:text-cyan-700"
                  >
                    📎
                  </button>
                  <span className="pointer-events-none absolute bottom-4 left-14 text-xs text-slate-400">
                    模板
                  </span>
                  <button
                    data-testid="extract-start"
                    onClick={handleExtract}
                    disabled={busy !== null}
                    className="absolute bottom-2 right-2 rounded-full bg-cyan-500 px-4 py-2 text-sm font-semibold text-white shadow-lg shadow-cyan-500/25 transition hover:bg-cyan-400 disabled:opacity-50"
                  >
                    {busy === "extract" ? "提取中" : "开始"}
                  </button>
                </div>
              </div>
              <input
                data-testid="extract-template-file-input"
                data-infowb-file-input="template"
                ref={templateInputRef}
                className="hidden"
                type="file"
                accept={`${SPREADSHEET_ACCEPT},image/*`}
                onChange={(event) => {
                  void handleTemplateFiles(event.target.files);
                  event.currentTarget.value = "";
                }}
              />

              <div className="mt-3 grid gap-2 text-sm">
                <div className="rounded-2xl bg-slate-100 px-3 py-2 text-slate-700">
                  Excel / CSV 模板：
                  {excelTemplate
                    ? excelTemplate.columns.join(" / ")
                    : "未上传，可选"}
                </div>
                <div className="rounded-2xl bg-slate-100 px-3 py-2 text-slate-700">
                  模板图片：
                  {screenshotTemplate
                    ? screenshotTemplate.name
                    : "未上传，可选"}
                </div>
              </div>

              {busy && busy !== "extract" ? (
                <p className="mt-3 text-sm text-cyan-700">
                  正在处理文件或模板…
                </p>
              ) : null}
              <p className="mt-3 text-sm text-slate-500">{status}</p>
              {!hasTemplate ? (
                <p className="mt-2 text-xs text-slate-400">
                  如果不上传模板，系统会根据文字要求自动整理表格列。
                </p>
              ) : null}
              {error ? (
                <div className="mt-3 rounded-2xl border border-red-200 bg-red-50 p-3 text-sm text-red-700">
                  {error}
                </div>
              ) : null}
            </div>
          </section>

          <section className="flex min-h-[560px] min-w-0 flex-col overflow-hidden rounded-3xl border border-white/10 bg-white text-slate-950 shadow-2xl shadow-black/20 lg:min-h-0">
            <div className="flex flex-col gap-3 border-b border-slate-200 bg-slate-50 p-4 md:flex-row md:items-center md:justify-between">
              <div>
                <p className="text-xs font-semibold uppercase tracking-[0.22em] text-cyan-700">
                  Excel Result
                </p>
                <h2 className="text-xl font-bold">Excel 结果</h2>
              </div>
              <div className="flex gap-2">
                <button
                  className="rounded-full bg-slate-900 px-4 py-2 text-sm font-semibold text-white disabled:opacity-40"
                  onClick={copyTsv}
                  disabled={!result}
                >
                  复制表格
                </button>
                <button
                  className="rounded-full bg-cyan-600 px-4 py-2 text-sm font-semibold text-white disabled:opacity-40"
                  onClick={downloadXlsx}
                  disabled={!result}
                >
                  下载 .xlsx
                </button>
              </div>
            </div>

            {result ? (
              <div className="min-h-0 flex-1 overflow-auto">
                <table data-testid="extract-result-table" className="w-full min-w-[720px] text-left text-sm">
                  <thead className="sticky top-0 z-10 bg-slate-100 shadow-sm">
                    <tr>
                      {result.columns.map((column) => (
                        <th
                          className="border-b border-slate-200 px-4 py-3 font-semibold text-slate-700"
                          key={column}
                        >
                          {column}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {result.rows.map((row, rowIndex) => (
                      <tr
                        className="odd:bg-white even:bg-slate-50/70"
                        key={`${rowIndex}-${row.join("-")}`}
                      >
                        {result.columns.map((column, columnIndex) => (
                          <td
                            className="border-b border-slate-100 px-4 py-3 align-top text-slate-700"
                            key={`${column}-${columnIndex}`}
                          >
                            {row[columnIndex]}
                          </td>
                        ))}
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            ) : (
              <div className="grid flex-1 place-items-center p-8 text-center">
                <div className="max-w-sm rounded-3xl bg-slate-50 p-6">
                  <p className="text-lg font-semibold text-slate-900">
                    结果会显示在这里
                  </p>
                  <p className="mt-2 text-sm leading-6 text-slate-500">
                    添加资料和提取要求后，点击开始提取。生成的表格可以直接复制或下载为
                    Excel。
                  </p>
                </div>
              </div>
            )}
          </section>
        </div>
      </div>
    </main>
  );
}
