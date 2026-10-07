import type {
  CompareCoveragePayload,
  CompareRequestPayload,
  ExtractedTable,
  SourceFilePayload,
  SourceTable,
} from "./types";
import { normalizeExtractedTable } from "./extraction";

export const COMPARE_COLUMNS = [
  "状态",
  "检查对象",
  "检查字段",
  "待核对值",
  "参考资料值",
  "结论",
  "置信度",
  "差异说明",
  "证据来源",
  "建议处理",
];

export const COMPARE_BATCH_SIZE = 20;

type CoverageStatus = "not_needed" | "filled_missing_rows";

export type CompareRowRecord = {
  rowId: string;
  fileName: string;
  sheetName: string;
  excelRowNumber: number;
  columns: string[];
  values: string[];
  rowValues: Record<string, string>;
  displayObject: string;
  serializedValues: string;
};

type CoverageCheck = {
  table: ExtractedTable;
  preFallbackMissingRowIds: string[];
  fallbackRowIds: string[];
};

type ReferenceEvidenceCheck = {
  table: ExtractedTable;
  missingEvidenceRowIds: string[];
};

function uniqueColumnName(
  columns: string[],
  index: number,
  usedColumns: Set<string>,
) {
  const baseName = columns[index]?.trim() || `Column ${index + 1}`;
  let columnName = baseName;
  let duplicateIndex = 2;
  while (usedColumns.has(columnName)) {
    columnName = `${baseName} ${duplicateIndex}`;
    duplicateIndex += 1;
  }
  usedColumns.add(columnName);
  return columnName;
}

function createRowValues(table: SourceTable, values: string[]) {
  const columnCount = Math.max(table.columns.length, values.length);
  const usedColumns = new Set<string>();
  const columns = Array.from({ length: columnCount }, (_, columnIndex) =>
    uniqueColumnName(table.columns, columnIndex, usedColumns),
  );
  const rowValues = columns.reduce<Record<string, string>>(
    (result, column, columnIndex) => {
      result[column] = values[columnIndex] ?? "";
      return result;
    },
    {},
  );
  return { columns, rowValues };
}

function summarizeRowValues(rowValues: Record<string, string>, maxItems = 5) {
  const entries = Object.entries(rowValues).filter(([, value]) => value.trim());
  const summary = entries
    .slice(0, maxItems)
    .map(([column, value]) => `${column}=${value}`)
    .join(" / ");
  return summary || "空行";
}

function createRowId(fileName: string, sheetName: string, excelRowNumber: number) {
  return `${fileName}#${sheetName}!R${excelRowNumber}`;
}

function tableRowsToRecords(
  file: SourceFilePayload,
  table: SourceTable,
): CompareRowRecord[] {
  return table.rows
    .map((values, rowIndex) => {
      const excelRowNumber = rowIndex + 2;
      const { columns, rowValues } = createRowValues(table, values);
      const rowId = createRowId(file.name, table.sheetName, excelRowNumber);
      return {
        rowId,
        fileName: file.name,
        sheetName: table.sheetName,
        excelRowNumber,
        columns,
        values: columns.map((_, columnIndex) => values[columnIndex] ?? ""),
        rowValues,
        displayObject: `${rowId} | ${summarizeRowValues(rowValues)}`,
        serializedValues: JSON.stringify(rowValues),
      };
    })
    .filter((record) => record.values.some((value) => value.trim()));
}

export function collectSubjectSpreadsheetRows(
  payload: CompareRequestPayload,
): CompareRowRecord[] {
  return payload.subject.files.flatMap((file) =>
    file.kind === "excel"
      ? (file.tables ?? []).flatMap((table) => tableRowsToRecords(file, table))
      : [],
  );
}

export function chunkCompareRows(
  rows: CompareRowRecord[],
  batchSize = COMPARE_BATCH_SIZE,
) {
  const chunks: CompareRowRecord[][] = [];
  for (let startIndex = 0; startIndex < rows.length; startIndex += batchSize) {
    chunks.push(rows.slice(startIndex, startIndex + batchSize));
  }
  return chunks;
}

function summarizeFiles(files: CompareRequestPayload["subject"]["files"]) {
  return files.map((file) => ({
    kind: file.kind,
    name: file.name,
    pageCount: file.pageCount,
    text: file.text?.slice(0, 30000) ?? "",
    tables: file.tables?.map((table) => ({
      sheetName: table.sheetName,
      columns: table.columns,
      rows: table.rows,
    })),
    imageCount: file.images?.length ?? 0,
  }));
}

function compactBatchRows(rows: CompareRowRecord[]) {
  return rows.map((record) => ({
    rowId: record.rowId,
    fileName: record.fileName,
    sheetName: record.sheetName,
    excelRowNumber: record.excelRowNumber,
    displayObject: record.displayObject,
    rowValues: record.rowValues,
  }));
}

function normalizeCompareRow(row: string[]) {
  return COMPARE_COLUMNS.map((_, columnIndex) => row[columnIndex] ?? "");
}

export function normalizeCompareTable(value: unknown): ExtractedTable {
  const normalized = normalizeExtractedTable(value);
  const columnSet = new Set(normalized.columns);
  const columns = COMPARE_COLUMNS.filter((column) => columnSet.has(column));
  const completeColumns = columns.length === COMPARE_COLUMNS.length;
  if (!completeColumns) {
    const rows = normalized.rows.map((row) =>
      COMPARE_COLUMNS.map((column) => {
        const index = normalized.columns.indexOf(column);
        return index >= 0 ? row[index] ?? "" : "";
      }),
    );
    return {
      text: normalized.text,
      columns: COMPARE_COLUMNS,
      rows,
    };
  }
  return {
    text: normalized.text,
    columns: COMPARE_COLUMNS,
    rows: normalized.rows.map((row) =>
      COMPARE_COLUMNS.map(
        (column) => row[normalized.columns.indexOf(column)] ?? "",
      ),
    ),
  };
}

export function parseCompareJson(rawText: string): ExtractedTable {
  const trimmed = rawText.trim();
  const fenced = trimmed.match(/```(?:json)?\s*([\s\S]*?)```/i)?.[1];
  const candidate = fenced ?? trimmed.match(/\{[\s\S]*\}/)?.[0] ?? trimmed;
  return normalizeCompareTable(JSON.parse(candidate));
}

export function buildComparePrompt(payload: CompareRequestPayload) {
  return [
    "你是一个严谨的信息核对助手。请对比左侧“待核对主体”和右侧“参考信息源”，输出结构化检查表。",
    "第一优先适配工程/BOM/图纸核对，但不要只限 BOM；合同、报价、采购清单、实验数据、截图资料等也要按用户核对要求处理。",
    '必须只返回 JSON，不要返回 markdown。JSON schema 必须是：{"text": string, "columns": string[], "rows": string[][]}。',
    `columns 必须严格等于：${COMPARE_COLUMNS.join(" | ")}`,
    "rows 中每一行长度必须和 columns 一致。默认一行一个字段检查项；遇到非结构化资料时允许一行概括性检查项。",
    "结论只能使用：一致、不一致、左侧缺失、右侧缺失、需人工确认。",
    "置信度使用：高、中、低。低置信度、找不到明确依据、OCR/图片不清晰或匹配不确定时，结论必须写“需人工确认”，不能强行判断一致。",
    "状态建议：一致用 ✅，不一致用 ❌，缺失或需人工确认用 ⚠️。",
    `快捷模板：${payload.templateName || "用户未选择快捷模板"}`,
    `核对要求：${payload.requirement || "用户未填写要求，请做通用资料核对，并优先标出明显一致、不一致、缺失和需人工确认的项目。"}`,
    "待核对主体直接文本：",
    payload.subject.text.slice(0, 30000) || "（无直接文本输入）",
    "待核对主体文件摘要 JSON：",
    JSON.stringify(summarizeFiles(payload.subject.files), null, 2),
    "参考信息源直接文本：",
    payload.reference.text.slice(0, 30000) || "（无直接文本输入）",
    "参考信息源文件摘要 JSON：",
    JSON.stringify(summarizeFiles(payload.reference.files), null, 2),
  ].join("\n\n");
}

export function buildCompareBatchPrompt(
  payload: CompareRequestPayload,
  rows: CompareRowRecord[],
  batchIndex: number,
  batchCount: number,
) {
  return [
    "你是一个严谨的信息核对助手。现在进入 Excel/CSV/XLS/XLSX 逐行核对批处理模式。",
    "目标：对本批待核对主体的每一条结构化行逐行核对右侧参考信息源，不允许抽样、跳行、概括性省略。",
    '必须只返回 JSON，不要返回 markdown。JSON schema 必须是：{"text": string, "columns": string[], "rows": string[][]}。',
    `columns 必须严格等于：${COMPARE_COLUMNS.join(" | ")}`,
    "rows 中每一行长度必须和 columns 一致。",
    "强制覆盖规则：本批 expectedRowIds 中的每个 rowId 都必须至少出现在一条输出行的“检查对象”单元格里，且必须原样保留 rowId。可以对同一个 rowId 输出多条字段核对行，但不能少于每个 rowId 一条。",
    "如果参考信息源找不到某个 rowId/主键/字段依据，仍必须为该 rowId 输出一条结果；结论写“右侧缺失”或“需人工确认”，置信度写“低”，不要静默省略。",
    "结论只能使用：一致、不一致、左侧缺失、右侧缺失、需人工确认。",
    "置信度使用：高、中、低。低置信度、依据不明确、OCR/图片不清晰或匹配不确定时，结论必须写“需人工确认”。",
    "状态建议：一致用 ✅，不一致用 ❌，缺失或需人工确认用 ⚠️。",
    `批次：${batchIndex + 1}/${batchCount}`,
    `快捷模板：${payload.templateName || "用户未选择快捷模板"}`,
    `核对要求：${payload.requirement || "用户未填写要求，请做通用资料核对，并优先标出明显一致、不一致、缺失和需人工确认的项目。"}`,
    "本批 expectedRowIds：",
    JSON.stringify(rows.map((record) => record.rowId), null, 2),
    "本批待核对主体结构化行 JSON：",
    JSON.stringify(compactBatchRows(rows), null, 2),
    "待核对主体补充直接文本：",
    payload.subject.text.slice(0, 12000) || "（无直接文本输入）",
    "参考信息源直接文本：",
    payload.reference.text.slice(0, 30000) || "（无直接文本输入）",
    "参考信息源文件摘要/索引 JSON：",
    JSON.stringify(summarizeFiles(payload.reference.files), null, 2),
  ].join("\n\n");
}

export function collectReferenceCorpus(payload: CompareRequestPayload) {
  return [
    payload.reference.text,
    ...payload.reference.files.flatMap((file) => [
      file.name,
      file.text ?? "",
      ...(file.tables ?? []).flatMap((table) => [
        table.sheetName,
        table.columns.join(" "),
        ...table.rows.map((row) => row.join(" ")),
      ]),
    ]),
  ]
    .join("\n")
    .toLowerCase();
}

export function findPrimaryValue(record: CompareRowRecord) {
  const entries = Object.entries(record.rowValues).filter(([, value]) =>
    value.trim(),
  );
  const preferred = entries.find(([column]) =>
    /part|item|料号|物料|零件|编号|编码|id/i.test(column),
  );
  return preferred?.[1] ?? entries[0]?.[1] ?? record.rowId;
}

export function mockCompareRows(
  payload: CompareRequestPayload,
  rows: CompareRowRecord[],
): ExtractedTable {
  const referenceName = payload.reference.files[0]?.name || "参考信息源";
  const referenceCorpus = collectReferenceCorpus(payload);
  return {
    text: [
      "Mock 逐行核对已完成。",
      `已按 rowId 覆盖 ${rows.length} 条待核对主体行。`,
      "线上配置真实模型 env 后会输出基于资料的逐行核对检查表。",
    ].join("\n"),
    columns: COMPARE_COLUMNS,
    rows: rows.map((record) => {
      const primaryValue = findPrimaryValue(record);
      const primaryFound = referenceCorpus.includes(primaryValue.toLowerCase());
      return [
        primaryFound ? "⚠️" : "⚠️",
        record.displayObject,
        "整行覆盖",
        record.serializedValues,
        primaryFound
          ? `参考资料中找到主键/线索 ${primaryValue}，字段级判断需真实模型复核。`
          : "未找到明确参考依据",
        primaryFound ? "需人工确认" : "右侧缺失",
        "低",
        primaryFound
          ? "Mock 模式只验证逐行覆盖，不替代真实字段级核对。"
          : "参考资料未命中该主体行主键，不能静默省略。",
        referenceName,
        primaryFound ? "使用真实模型或人工确认字段差异。" : "补充参考资料或人工确认。",
      ];
    }),
  };
}

export function mockCompare(payload: CompareRequestPayload): ExtractedTable {
  const subjectRows = collectSubjectSpreadsheetRows(payload);
  if (subjectRows.length) return mockCompareRows(payload, subjectRows);

  const subjectName = payload.subject.files[0]?.name || "待核对主体";
  const referenceName = payload.reference.files[0]?.name || "参考信息源";
  const subjectText = [
    payload.subject.text,
    ...payload.subject.files.map((file) => file.text ?? file.name),
  ].join("\n");
  const referenceText = [
    payload.reference.text,
    ...payload.reference.files.map((file) => file.text ?? file.name),
  ].join("\n");

  const rows = [
    [
      "✅",
      "示例对象 A-001",
      "零件号/主键",
      subjectText.match(/[A-Z]-?\d{3,}/i)?.[0] || "A-001",
      referenceText.match(/[A-Z]-?\d{3,}/i)?.[0] || "A-001",
      "一致",
      "高",
      "两侧均找到相同主键或示例字段。",
      referenceName,
      "无需处理",
    ],
    [
      "❌",
      "示例对象 A-001",
      "版本/金额/关键值",
      "Rev.B",
      "Rev.C",
      "不一致",
      "中",
      "Mock 演示差异行；真实接口会根据上传资料生成。",
      referenceName,
      "请确认以参考资料还是主体资料为准。",
    ],
    [
      "⚠️",
      subjectName,
      "依据完整性",
      "已提供",
      payload.reference.files.length || payload.reference.text.trim()
        ? "部分资料可用"
        : "未找到明确依据",
      payload.reference.files.length || payload.reference.text.trim()
        ? "需人工确认"
        : "右侧缺失",
      "低",
      "依据不够明确，不能强行判定一致。",
      payload.reference.files.length ? referenceName : "-",
      "补充更清晰的参考资料后复核。",
    ],
  ];

  return {
    text: [
      "Mock 核对已完成。",
      `识别到待核对主体 ${payload.subject.files.length} 个文件、参考信息源 ${payload.reference.files.length} 个文件。`,
      "线上配置真实模型 env 后会输出基于资料的核对检查表。",
    ].join("\n"),
    columns: COMPARE_COLUMNS,
    rows,
  };
}

function fallbackCoverageRow(record: CompareRowRecord): string[] {
  return [
    "⚠️",
    record.displayObject,
    "整行覆盖",
    record.serializedValues,
    "模型未返回该 rowId 的逐行核对结果",
    "需人工确认",
    "低",
    "覆盖校验器发现模型输出遗漏该主体行，已补入兜底行以避免静默遗漏。",
    `${record.fileName} / ${record.sheetName} / Excel row ${record.excelRowNumber}`,
    "请人工确认或重新核对该行。",
  ];
}

function missingReferenceEvidenceRow(
  record: CompareRowRecord,
  primaryValue: string,
): string[] {
  return [
    "⚠️",
    record.displayObject,
    "参考依据覆盖",
    record.serializedValues,
    `参考资料未命中该行主键/线索：${primaryValue}`,
    "右侧缺失",
    "低",
    "覆盖校验器确认右侧参考资料没有该主体行的主键/线索，已保留明确缺证据结果，避免抽样遗漏。",
    `${record.fileName} / ${record.sheetName} / Excel row ${record.excelRowNumber}`,
    "补充参考资料或人工确认该行。",
  ];
}

function rowContainsRowId(row: string[], rowId: string) {
  const objectColumnIndex = COMPARE_COLUMNS.indexOf("检查对象");
  return String(row[objectColumnIndex] ?? "").includes(rowId);
}

function coveredRowIds(table: ExtractedTable, rows: CompareRowRecord[]) {
  return rows
    .filter((record) =>
      table.rows.some((row) => rowContainsRowId(normalizeCompareRow(row), record.rowId)),
    )
    .map((record) => record.rowId);
}

export function ensureRowCoverage(
  table: ExtractedTable,
  expectedRows: CompareRowRecord[],
): CoverageCheck {
  const normalizedRows = table.rows.map(normalizeCompareRow);
  const coveredIds = new Set(coveredRowIds({ ...table, rows: normalizedRows }, expectedRows));
  const missingRows = expectedRows.filter((record) => !coveredIds.has(record.rowId));
  const fallbackRows = missingRows.map(fallbackCoverageRow);
  return {
    table: {
      text: table.text,
      columns: COMPARE_COLUMNS,
      rows: [...normalizedRows, ...fallbackRows],
    },
    preFallbackMissingRowIds: missingRows.map((record) => record.rowId),
    fallbackRowIds: missingRows.map((record) => record.rowId),
  };
}

export function ensureReferenceEvidenceCoverage(
  table: ExtractedTable,
  expectedRows: CompareRowRecord[],
  referenceCorpus: string,
): ReferenceEvidenceCheck {
  const normalizedRows = table.rows.map(normalizeCompareRow);
  const missingEvidenceRows = expectedRows.filter((record) => {
    const primaryValue = findPrimaryValue(record).trim().toLowerCase();
    return Boolean(primaryValue) && !referenceCorpus.includes(primaryValue);
  });
  const rowsToAppend = missingEvidenceRows.filter((record) => {
    const objectColumnIndex = COMPARE_COLUMNS.indexOf("检查对象");
    const conclusionColumnIndex = COMPARE_COLUMNS.indexOf("结论");
    return !normalizedRows.some((row) => {
      const hasRowId = String(row[objectColumnIndex] ?? "").includes(record.rowId);
      const conclusion = String(row[conclusionColumnIndex] ?? "");
      return hasRowId && ["右侧缺失", "需人工确认"].includes(conclusion);
    });
  });
  return {
    table: {
      text: table.text,
      columns: COMPARE_COLUMNS,
      rows: [
        ...normalizedRows,
        ...rowsToAppend.map((record) =>
          missingReferenceEvidenceRow(record, findPrimaryValue(record)),
        ),
      ],
    },
    missingEvidenceRowIds: missingEvidenceRows.map((record) => record.rowId),
  };
}

export function mergeCompareTables(tables: ExtractedTable[]): ExtractedTable {
  const seenRows = new Set<string>();
  const mergedRows: string[][] = [];
  for (const table of tables) {
    for (const row of table.rows.map(normalizeCompareRow)) {
      const rowKey = JSON.stringify(row);
      if (seenRows.has(rowKey)) continue;
      seenRows.add(rowKey);
      mergedRows.push(row);
    }
  }
  return {
    text: tables.map((table) => table.text).filter(Boolean).join("\n\n"),
    columns: COMPARE_COLUMNS,
    rows: mergedRows,
  };
}

export function createCompareCoverageReport(
  table: ExtractedTable,
  expectedRows: CompareRowRecord[],
  options: {
    batchCount: number;
    repairCount: number;
    rerunCount: number;
    fallbackRowIds: string[];
    preFallbackMissingRowIds: string[];
    missingEvidenceRowIds: string[];
    batchSize: number;
  },
): CompareCoveragePayload {
  const outputRowIds = coveredRowIds(table, expectedRows);
  const outputRowIdSet = new Set(outputRowIds);
  const expectedRowIds = expectedRows.map((record) => record.rowId);
  const missingRowIds = expectedRowIds.filter(
    (rowId) => !outputRowIdSet.has(rowId),
  );
  const fallbackRowIds = Array.from(new Set(options.fallbackRowIds));
  const preFallbackMissingRowIds = Array.from(
    new Set(options.preFallbackMissingRowIds),
  );
  const fallbackStatus: CoverageStatus = fallbackRowIds.length
    ? "filled_missing_rows"
    : "not_needed";
  const missingEvidenceRowIds = Array.from(new Set(options.missingEvidenceRowIds));
  return {
    inputRowCount: expectedRows.length,
    expectedRowIds,
    outputRowIds,
    coveredRowIds: outputRowIds,
    missingRowIds,
    missingRowCount: missingRowIds.length,
    preFallbackMissingRowIds,
    preFallbackMissingRowCount: preFallbackMissingRowIds.length,
    fallbackRowIds,
    fallbackCount: fallbackRowIds.length,
    fallbackStatus,
    missingEvidenceRowIds,
    missingEvidenceRowCount: missingEvidenceRowIds.length,
    batchCount: options.batchCount,
    batchSize: options.batchSize,
    repairCount: options.repairCount,
    rerunCount: options.rerunCount,
  };
}

export function mockCompareRepair(
  rawText: string,
  payload: CompareRequestPayload,
  rows?: CompareRowRecord[],
): ExtractedTable {
  const repaired = rows?.length ? mockCompareRows(payload, rows) : mockCompare(payload);
  return {
    ...repaired,
    text: `${repaired.text}\n\n已触发核对 JSON 修复路径；原始坏输出长度 ${rawText.length}。`,
  };
}
