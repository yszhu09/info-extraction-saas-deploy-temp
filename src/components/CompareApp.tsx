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
  CompareRequestPayload,
  CompareResponsePayload,
  SourceFilePayload,
  SourceImage,
} from "@/lib/types";

const MAX_MODEL_IMAGES = 8;
const MAX_IMAGE_DIMENSION = 1400;
const IMAGE_JPEG_QUALITY = 0.72;
const MIN_IMAGE_DIMENSION = 16;

type BusyState = "subject" | "reference" | "compare" | null;
type Side = "subject" | "reference";

const sourceKindLabel: Record<SourceFilePayload["kind"], string> = {
  pdf: "PDF",
  image: "图片",
  excel: "Excel",
};

const comparisonTemplates = [
  {
    name: "BOM/图纸核对",
    prompt:
      "检查零件号、版本、数量、规格、材料等是否一致；如果参考资料中缺失、图纸与 BOM 不一致、或依据不清晰，都要标出并给出证据来源。",
  },
  {
    name: "合同信息核对",
    prompt:
      "检查合同金额、日期、客户名称、项目名称、付款节点、关键条款等是否一致；金额或日期不明确时标为需人工确认。",
  },
  {
    name: "实验数据核对",
    prompt:
      "检查样品编号、测试项目、结果值、单位、判定标准和结论是否一致；数据缺失或单位不一致时单独列出。",
  },
  {
    name: "通用资料核对",
    prompt:
      "根据左侧主体和右侧参考资料，找出关键字段的一致、不一致、缺失和需要人工确认的项目；不要编造资料中没有的依据。",
  },
];

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

async function processCompareFile(file: File): Promise<SourceFilePayload> {
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

function countSideImages(files: SourceFilePayload[]) {
  return files.reduce((count, file) => count + (file.images?.length ?? 0), 0);
}

function fileSummary(file: SourceFilePayload) {
  return `${sourceKindLabel[file.kind]}${file.pageCount ? ` · ${file.pageCount} 页` : ""}`;
}

export function CompareApp() {
  const [subjectText, setSubjectText] = useState("");
  const [subjectFiles, setSubjectFiles] = useState<SourceFilePayload[]>([]);
  const [referenceText, setReferenceText] = useState("");
  const [referenceFiles, setReferenceFiles] = useState<SourceFilePayload[]>([]);
  const [requirement, setRequirement] = useState(comparisonTemplates[0].prompt);
  const [templateName, setTemplateName] = useState(comparisonTemplates[0].name);
  const [result, setResult] = useState<CompareResponsePayload | null>(null);
  const [busy, setBusy] = useState<BusyState>(null);
  const [status, setStatus] = useState("请先添加待核对主体和参考信息源。");
  const [error, setError] = useState("");
  const subjectInputRef = useRef<HTMLInputElement>(null);
  const referenceInputRef = useRef<HTMLInputElement>(null);
  const subjectTextAreaRef = useRef<HTMLTextAreaElement>(null);
  const referenceTextAreaRef = useRef<HTMLTextAreaElement>(null);

  const hasSubject = subjectText.trim().length > 0 || subjectFiles.length > 0;
  const hasReference =
    referenceText.trim().length > 0 || referenceFiles.length > 0;

  function setFilesForSide(
    side: Side,
    updater: (current: SourceFilePayload[]) => SourceFilePayload[],
  ) {
    if (side === "subject") setSubjectFiles(updater);
    else setReferenceFiles(updater);
  }

  async function handleFiles(side: Side, files: FileList | File[] | null) {
    const nextFiles = Array.from(files ?? []);
    if (!nextFiles.length) return;
    setBusy(side);
    setError("");
    try {
      const processed: SourceFilePayload[] = [];
      const skipped: string[] = [];
      for (const file of nextFiles) {
        try {
          processed.push(await processCompareFile(file));
        } catch (fileError) {
          skipped.push(
            fileError instanceof Error
              ? fileError.message
              : `${file.name} 处理失败`,
          );
        }
      }
      if (processed.length) {
        setFilesForSide(side, (current) => [...current, ...processed]);
        setStatus(
          `${side === "subject" ? "待核对主体" : "参考信息源"}已添加 ${processed.length} 个文件；图片会自动压缩，核对时最多发送前 ${MAX_MODEL_IMAGES} 张图片。`,
        );
      }
      if (skipped.length) {
        setError(`部分文件未添加：${skipped.slice(0, 3).join("；")}`);
      }
      if (!processed.length && !skipped.length) {
        setStatus("未检测到可处理文件。");
      }
    } catch (fileError) {
      setError(fileError instanceof Error ? fileError.message : "文件处理失败");
    } finally {
      setBusy(null);
    }
  }

  function stopDrag(event: DragEvent<HTMLElement>) {
    event.preventDefault();
    event.stopPropagation();
  }

  async function handleDrop(side: Side, event: DragEvent<HTMLElement>) {
    stopDrag(event);
    await handleFiles(side, event.dataTransfer.files);
  }

  async function handlePaste(side: Side, event: ClipboardEvent<HTMLElement>) {
    const images = clipboardImages(event);
    if (!images.length) return;
    event.preventDefault();
    await handleFiles(side, images);
  }

  useEffect(() => {
    const subjectElement = subjectTextAreaRef.current;
    const referenceElement = referenceTextAreaRef.current;

    const handleSubjectTextAreaPaste = (event: globalThis.ClipboardEvent) => {
      const images = clipboardImagesFromData(event.clipboardData);
      if (!images.length) return;
      event.preventDefault();
      event.stopPropagation();
      void handleFiles("subject", images);
    };
    const handleReferenceTextAreaPaste = (event: globalThis.ClipboardEvent) => {
      const images = clipboardImagesFromData(event.clipboardData);
      if (!images.length) return;
      event.preventDefault();
      event.stopPropagation();
      void handleFiles("reference", images);
    };

    subjectElement?.addEventListener("paste", handleSubjectTextAreaPaste);
    referenceElement?.addEventListener("paste", handleReferenceTextAreaPaste);

    return () => {
      subjectElement?.removeEventListener("paste", handleSubjectTextAreaPaste);
      referenceElement?.removeEventListener(
        "paste",
        handleReferenceTextAreaPaste,
      );
    };
  });

  function applyTemplate(template: (typeof comparisonTemplates)[number]) {
    setTemplateName(template.name);
    setRequirement(template.prompt);
    setStatus(`已套用「${template.name}」，可以继续编辑核对要求。`);
  }

  async function handleCompare() {
    if (!hasSubject) {
      setError("请先输入或上传待核对主体。");
      return;
    }
    if (!hasReference) {
      setError("请先输入或上传参考信息源。");
      return;
    }
    setBusy("compare");
    setError("");
    setResult(null);
    try {
      const payload: CompareRequestPayload = {
        subject: { text: subjectText, files: subjectFiles },
        reference: { text: referenceText, files: referenceFiles },
        requirement,
        templateName,
      };
      const imageCount =
        countSideImages(subjectFiles) + countSideImages(referenceFiles);
      const imageStatus = imageCount
        ? imageCount > MAX_MODEL_IMAGES
          ? `图片已压缩；本次优先发送前 ${MAX_MODEL_IMAGES} 张，超出的仅使用文字摘要。`
          : "图片已压缩后发送给模型。"
        : "";
      setStatus(imageStatus || "正在核对并生成检查表…");
      const response = await fetch("/api/compare", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload),
      });
      const body = await response.json();
      if (!response.ok) throw new Error(body.error || "核对失败");
      setResult(body as CompareResponsePayload);
      setStatus(
        imageStatus
          ? `核对完成，已生成检查表。${imageStatus}`
          : "核对完成，已生成检查表。",
      );
    } catch (compareError) {
      setError(compareError instanceof Error ? compareError.message : "核对失败");
    } finally {
      setBusy(null);
    }
  }

  async function copyTsv() {
    if (!result) return;
    await navigator.clipboard.writeText(toTsv(result.columns, result.rows));
    setStatus("检查表已复制到剪贴板。");
  }

  async function downloadXlsx() {
    if (!result) return;
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet("Comparison Check");
    sheet.addRow(result.columns);
    result.rows.forEach((row) => sheet.addRow(row));
    sheet.columns.forEach((column) => {
      column.width = Math.min(
        42,
        Math.max(
          12,
          ...((column.values ?? []).map((value) => String(value ?? "").length) as
            | number[]
            | []),
        ),
      );
    });
    const buffer = await workbook.xlsx.writeBuffer();
    const blob = new Blob([bufferToArrayBuffer(buffer)], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = "comparison-check.xlsx";
    link.click();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  async function logout() {
    await fetch("/api/logout", { method: "POST" });
    window.location.href = "/login";
  }

  function renderMessageBox(side: Side) {
    const isSubject = side === "subject";
    const title = isSubject ? "待核对主体" : "参考信息源";
    const label = isSubject ? "Subject" : "Reference";
    const helper = isSubject
      ? "要检查的主文件/主内容，例如 BOM、清单、合同、报价、实验数据。"
      : "用来作为依据的资料，例如图纸、设计资料、PDF、图片、Excel、文本。";
    const value = isSubject ? subjectText : referenceText;
    const files = isSubject ? subjectFiles : referenceFiles;
    const textAreaRef = isSubject ? subjectTextAreaRef : referenceTextAreaRef;
    const inputRef = isSubject ? subjectInputRef : referenceInputRef;
    const setText = isSubject ? setSubjectText : setReferenceText;
    const busyForSide = busy === side;

    return (
      <section className="rounded-3xl bg-white p-4 text-slate-950 shadow-xl">
        <div className="flex items-start justify-between gap-3">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.22em] text-cyan-700">
              {label}
            </p>
            <h2 className="text-xl font-bold">{title}</h2>
            <p className="mt-1 text-sm leading-6 text-slate-500">{helper}</p>
          </div>
          {files.length ? (
            <button
              className="rounded-full bg-slate-100 px-3 py-1 text-xs text-slate-600 hover:bg-slate-200"
              onClick={() => setFilesForSide(side, () => [])}
            >
              清空文件
            </button>
          ) : null}
        </div>

        <div
          data-testid={`${side}-dropzone`}
          data-infowb-dropzone={side}
          onDragEnter={stopDrag}
          onDragOver={stopDrag}
          onDrop={(event) => void handleDrop(side, event)}
          onPaste={(event) => void handlePaste(side, event)}
          className="mt-4 rounded-3xl border border-slate-200 bg-slate-50 p-3 transition focus-within:border-cyan-500 focus-within:bg-white"
        >
          <div className="relative">
            <textarea
              data-testid={`${side}-textarea`}
              data-infowb-textarea={side}
              ref={textAreaRef}
              className="min-h-52 w-full resize-none rounded-2xl bg-transparent p-3 pb-16 text-sm outline-none placeholder:text-slate-400"
              value={value}
              onChange={(event) => setText(event.target.value)}
              placeholder={
                isSubject
                  ? "输入待核对内容，也可以粘贴截图、拖入文件，或点回形针上传 Excel / CSV / PDF / 图片"
                  : "输入参考依据，也可以粘贴图纸截图、拖入资料，或点回形针上传 PDF / 图片 / Excel / CSV"
              }
            />
            <button
              type="button"
              aria-label={`上传${title}文件`}
              onClick={() => inputRef.current?.click()}
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
          data-testid={`${side}-file-input`}
          data-infowb-file-input={side}
          ref={inputRef}
          className="hidden"
          type="file"
          multiple
          accept={`application/pdf,image/*,${SPREADSHEET_ACCEPT}`}
          onChange={(event) => {
            void handleFiles(side, event.target.files);
            event.currentTarget.value = "";
          }}
        />

        <div className="mt-4 space-y-2">
          {files.length ? (
            files.map((file, index) => (
              <div
                className="flex items-center justify-between gap-3 rounded-2xl bg-slate-100 px-3 py-2 text-sm text-slate-700"
                key={`${side}-${file.name}-${index}`}
              >
                <span className="min-w-0 truncate">
                  {file.name}
                  <span className="ml-2 text-xs text-slate-500">
                    {fileSummary(file)}
                  </span>
                </span>
                <button
                  className="shrink-0 text-xs text-slate-500 hover:text-red-600"
                  onClick={() =>
                    setFilesForSide(side, (current) =>
                      current.filter((_, itemIndex) => itemIndex !== index),
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
        {busyForSide ? (
          <p className="mt-3 text-sm text-cyan-700">正在处理文件…</p>
        ) : null}
      </section>
    );
  }

  return (
    <main className="min-h-screen bg-slate-950 text-slate-100" data-infowb-page="compare">
      <div className="mx-auto flex max-w-7xl flex-col gap-5 px-4 py-5 lg:min-h-screen lg:px-6">
        <header className="flex flex-col gap-3 rounded-3xl border border-white/10 bg-white/8 px-5 py-4 shadow-2xl shadow-black/20 md:flex-row md:items-center md:justify-between">
          <div>
            <p className="text-xs font-semibold uppercase tracking-[0.28em] text-cyan-300">
              Compare Workbench
            </p>
            <h1 className="mt-1 text-2xl font-bold tracking-tight md:text-3xl">
              信息核对工作台
            </h1>
          </div>
          <div className="flex flex-wrap items-center gap-3 text-sm text-slate-300">
            <nav className="flex rounded-full border border-white/15 bg-slate-900/60 p-1">
              <Link
                prefetch={false}
                href="/"
                className="rounded-full px-3 py-1.5 text-slate-300 hover:text-cyan-200"
              >
                信息提取
              </Link>
              <Link
                prefetch={false}
                href="/compare"
                className="rounded-full bg-cyan-400 px-3 py-1.5 font-semibold text-slate-950"
              >
                信息核对
              </Link>
            </nav>
            <JeStatusBadge page="compare" />
            <span className="rounded-full border border-cyan-300/30 bg-cyan-300/10 px-3 py-1 text-cyan-100">
              主体资料 → 参考依据 → Excel 检查表
            </span>
            <button
              onClick={logout}
              className="rounded-full border border-white/20 px-4 py-2 text-slate-200 hover:border-cyan-300 hover:text-cyan-200"
            >
              退出
            </button>
          </div>
        </header>

        <section className="rounded-3xl border border-cyan-300/20 bg-cyan-300/10 p-4 text-sm leading-6 text-cyan-50">
          左侧放要核对的主体，右侧放作为依据的信息源。下方选择或编辑核对要求后，AI 会生成一致、不一致、缺失和需人工确认的检查表。
        </section>

        <div className="grid gap-5 xl:grid-cols-2">
          {renderMessageBox("subject")}
          {renderMessageBox("reference")}
        </div>

        <section className="rounded-3xl bg-white p-4 text-slate-950 shadow-xl">
          <div className="flex flex-col gap-3 lg:flex-row lg:items-start lg:justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.22em] text-cyan-700">
                Check Requirement
              </p>
              <h2 className="text-xl font-bold">核对要求 / 自定义检查字段</h2>
              <p className="mt-1 text-sm text-slate-500">
                可先点快捷模板，再按这次实际资料修改要检查的字段。
              </p>
            </div>
            <div className="flex flex-wrap gap-2">
              {comparisonTemplates.map((template) => (
                <button
                  data-testid={`template-${template.name}`}
                  key={template.name}
                  onClick={() => applyTemplate(template)}
                  className={`rounded-full border px-3 py-1.5 text-sm transition ${
                    templateName === template.name
                      ? "border-cyan-500 bg-cyan-50 text-cyan-700"
                      : "border-slate-200 bg-slate-50 text-slate-600 hover:border-cyan-300"
                  }`}
                >
                  {template.name}
                </button>
              ))}
            </div>
          </div>
          <div className="mt-4 rounded-3xl border border-slate-200 bg-slate-50 p-3 transition focus-within:border-cyan-500 focus-within:bg-white">
            <div className="relative">
              <textarea
                data-testid="compare-requirement"
                className="min-h-32 w-full resize-none rounded-2xl bg-transparent p-3 pb-16 text-sm outline-none placeholder:text-slate-400"
                value={requirement}
                onChange={(event) => setRequirement(event.target.value)}
                placeholder="例如：检查零件号、版本、数量是否一致；资料缺失、低置信度或依据不明确时标为需人工确认。"
              />
              <button
                data-testid="compare-start"
                onClick={handleCompare}
                disabled={busy !== null}
                className="absolute bottom-2 right-2 rounded-full bg-cyan-500 px-5 py-2 text-sm font-semibold text-white shadow-lg shadow-cyan-500/25 transition hover:bg-cyan-400 disabled:opacity-50"
              >
                {busy === "compare" ? "核对中" : "开始核对"}
              </button>
            </div>
          </div>
          <p className="mt-3 text-sm text-slate-500">{status}</p>
          {error ? (
            <div className="mt-3 rounded-2xl border border-red-200 bg-red-50 p-3 text-sm text-red-700">
              {error}
            </div>
          ) : null}
        </section>

        <section className="flex min-h-[520px] min-w-0 flex-col overflow-hidden rounded-3xl border border-white/10 bg-white text-slate-950 shadow-2xl shadow-black/20">
          <div className="flex flex-col gap-3 border-b border-slate-200 bg-slate-50 p-4 md:flex-row md:items-center md:justify-between">
            <div>
              <p className="text-xs font-semibold uppercase tracking-[0.22em] text-cyan-700">
                Comparison Check
              </p>
              <h2 className="text-xl font-bold">Excel 检查表</h2>
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
              <table data-testid="compare-result-table" className="w-full min-w-[1100px] text-left text-sm">
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
              <div className="max-w-md rounded-3xl bg-slate-50 p-6">
                <p className="text-lg font-semibold text-slate-900">
                  检查表会显示在这里
                </p>
                <p className="mt-2 text-sm leading-6 text-slate-500">
                  添加待核对主体和参考信息源后，选择或编辑核对要求，点击开始核对。生成的检查表可以复制或下载为 Excel。
                </p>
              </div>
            </div>
          )}
        </section>
      </div>
    </main>
  );
}
