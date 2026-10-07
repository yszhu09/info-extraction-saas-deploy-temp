import type { ExtractedTable, ExtractRequestPayload } from "./types";

export function normalizeExtractedTable(value: unknown): ExtractedTable {
  if (!value || typeof value !== "object") {
    throw new Error("AI 输出不是对象。");
  }
  const record = value as Record<string, unknown>;
  const text = typeof record.text === "string" ? record.text : "";
  const columns = Array.isArray(record.columns)
    ? record.columns.map((item) => String(item || "").trim()).filter(Boolean)
    : [];
  const rawRows = Array.isArray(record.rows) ? record.rows : [];
  const rows = rawRows.map((row) =>
    Array.isArray(row)
      ? columns.map((_, index) => String(row[index] ?? ""))
      : columns.map((column) =>
          String((row as Record<string, unknown>)?.[column] ?? ""),
        ),
  );

  if (columns.length === 0) throw new Error("AI 输出缺少 columns。");

  return {
    text: text.trim() || "已生成结构化表格。",
    columns,
    rows,
  };
}

export function parseExtractedJson(rawText: string): ExtractedTable {
  const trimmed = rawText.trim();
  const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i)?.[1];
  const candidate = fenced ?? trimmed.match(/\{[\s\S]*\}/)?.[0] ?? trimmed;
  return normalizeExtractedTable(JSON.parse(candidate));
}

export function buildExtractionPrompt(payload: ExtractRequestPayload) {
  const templateColumns = payload.template.excelTemplate?.columns ?? [];
  const fileSummaries = payload.files.map((file) => ({
    kind: file.kind,
    name: file.name,
    pageCount: file.pageCount,
    text: file.text?.slice(0, 12000) ?? "",
    tables: file.tables?.map((table) => ({
      sheetName: table.sheetName,
      columns: table.columns,
      rows: table.rows.slice(0, 12),
    })),
    imageCount: file.images?.length ?? 0,
  }));

  return [
    "你是一个信息提取器。请根据资料源和目标格式约束，抽取用户需要的信息。",
    '必须只返回 JSON，不要返回 markdown。JSON schema 必须是：{"text": string, "columns": string[], "rows": string[][]}。',
    "rows 中每一行长度必须和 columns 一致；无法确认的单元格填空字符串，不要编造。",
    `输出偏好：${payload.template.outputMode}`,
    `文字要求：${payload.template.instruction || "用户未填写额外文字要求，请按资料源提取关键信息。"}`,
    templateColumns.length
      ? `Excel 模板字段：${templateColumns.join(" | ")}`
      : "Excel 模板字段：未提供。",
    payload.template.screenshotTemplate
      ? "用户提供了截图模板，请参考截图中的表头/版式。"
      : "截图模板：未提供。",
    "资料源文本：",
    payload.sourceText.slice(0, 16000) || "（无直接文本输入）",
    "资料源文件摘要 JSON：",
    JSON.stringify(fileSummaries, null, 2),
  ].join("\n\n");
}

export function mockExtract(payload: ExtractRequestPayload): ExtractedTable {
  const templateColumns =
    payload.template.excelTemplate?.columns.filter(Boolean) ?? [];
  const columns =
    templateColumns.length >= 2 ? templateColumns : ["字段", "提取值", "来源"];
  const rows: string[][] = [];

  const textLines = payload.sourceText
    .split(/\n+/)
    .map((line) => line.trim())
    .filter(Boolean)
    .slice(0, 4);

  if (textLines.length) {
    textLines.forEach((line, index) => {
      rows.push(
        columns.map((column, columnIndex) => {
          if (columnIndex === 0)
            return column.includes("字段")
              ? `文本片段 ${index + 1}`
              : line.slice(0, 80);
          if (columnIndex === 1) return line.slice(0, 120);
          return "文本输入";
        }),
      );
    });
  }

  payload.files.forEach((file) => {
    const source =
      file.kind === "pdf"
        ? `PDF ${file.pageCount ?? "?"}页`
        : file.kind === "excel"
          ? "Excel"
          : "图片";
    rows.push(
      columns.map((column, columnIndex) => {
        if (columnIndex === 0) return file.name;
        if (columnIndex === 1)
          return file.text?.slice(0, 120) || `${source} 已接入 mock 提取`;
        return source;
      }),
    );
  });

  if (payload.template.screenshotTemplate) {
    rows.push(
      columns.map((_, index) =>
        index === 0
          ? "截图模板"
          : index === 1
            ? (payload.template.screenshotTemplate?.name ?? "已提供")
            : "模板约束",
      ),
    );
  }

  if (rows.length === 0) {
    rows.push(
      columns.map((_, index) =>
        index === 0
          ? "示例字段"
          : index === 1
            ? "请提供资料源后重新提取"
            : "mock",
      ),
    );
  }

  return {
    text: [
      "Mock 提取已完成。",
      `识别到 ${payload.files.length} 个文件、${payload.template.excelTemplate ? "1 个 Excel 模板" : "无 Excel 模板"}、${payload.template.screenshotTemplate ? "1 个截图模板" : "无截图模板"}。`,
      "配置 CHATGPT_API_URL / CHATGPT_API_KEY / CHATGPT_MODEL 后，后端会切换到真实模型。",
    ].join("\n"),
    columns,
    rows,
  };
}

export function mockRepair(
  rawText: string,
  payload: ExtractRequestPayload,
): ExtractedTable {
  const repaired = mockExtract(payload);
  return {
    ...repaired,
    text: `${repaired.text}\n\n已触发 JSON 修复路径；原始坏输出长度 ${rawText.length}。`,
  };
}
