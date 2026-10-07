import { cookies } from "next/headers";
import { NextResponse } from "next/server";
import {
  COMPARE_BATCH_SIZE,
  buildCompareBatchPrompt,
  buildComparePrompt,
  chunkCompareRows,
  collectReferenceCorpus,
  collectSubjectSpreadsheetRows,
  createCompareCoverageReport,
  ensureReferenceEvidenceCoverage,
  ensureRowCoverage,
  mergeCompareTables,
  mockCompare,
  mockCompareRepair,
  mockCompareRows,
  parseCompareJson,
} from "@/lib/compare";
import type {
  CompareCoveragePayload,
  CompareRequestPayload,
  ExtractedTable,
  SourceImage,
} from "@/lib/types";

export const runtime = "nodejs";

const MAX_MODEL_IMAGES = 8;
const MIN_IMAGE_DATA_URL_LENGTH = 1_500;
const MAX_IMAGE_DATA_URL_LENGTH = 6_500_000;
const MAX_BATCH_RERUNS = 1;

const SUPPORTED_MODEL_IMAGE_PREFIXES = [
  "data:image/jpeg;base64,",
  "data:image/png;base64,",
  "data:image/webp;base64,",
];

type AiContent = {
  type: "input_text" | "input_image";
  text?: string;
  image_url?: string;
};

type ModelMode = "mock" | "real";

type PromptRequest = {
  payload: CompareRequestPayload;
  prompt: string;
  includeImages?: boolean;
};

type ParsedModelOutput = {
  table: ExtractedTable;
  repaired: boolean;
  mode: ModelMode;
};

type SpreadsheetCompareResult = ParsedModelOutput & {
  coverage: CompareCoveragePayload;
};

function extractOutputText(value: unknown): string {
  if (typeof value === "string") return value;
  if (!value || typeof value !== "object") return "";
  const record = value as Record<string, unknown>;
  if (typeof record.output_text === "string") return record.output_text;
  if (typeof record.text === "string") return record.text;
  if (Array.isArray(record.output)) {
    return record.output
      .flatMap((item) => {
        const content = (item as Record<string, unknown>)?.content;
        if (!Array.isArray(content)) return [];
        return content.map((part) => {
          const partRecord = part as Record<string, unknown>;
          return partRecord.text ?? partRecord.output_text ?? "";
        });
      })
      .join("\n");
  }
  if (Array.isArray(record.choices)) {
    return record.choices
      .map((choice) => {
        const choiceRecord = choice as Record<string, unknown>;
        const message = choiceRecord.message as
          | Record<string, unknown>
          | undefined;
        return message?.content ?? choiceRecord.text ?? "";
      })
      .join("\n");
  }
  return "";
}

function parseSseText(rawText: string) {
  return rawText
    .split("\n")
    .filter((line) => line.startsWith("data:"))
    .map((line) => line.replace(/^data:\s*/, "").trim())
    .filter((line) => line && line !== "[DONE]")
    .map((line) => {
      try {
        return extractOutputText(JSON.parse(line));
      } catch {
        return line;
      }
    })
    .join("\n");
}

function isSupportedModelImage(dataUrl: string) {
  if (dataUrl.length < MIN_IMAGE_DATA_URL_LENGTH) return false;
  if (dataUrl.length > MAX_IMAGE_DATA_URL_LENGTH) return false;
  return SUPPORTED_MODEL_IMAGE_PREFIXES.some((prefix) =>
    dataUrl.startsWith(prefix),
  );
}

function collectImages(payload: CompareRequestPayload): SourceImage[] {
  return [
    ...payload.subject.files.flatMap((file) => file.images ?? []),
    ...payload.reference.files.flatMap((file) => file.images ?? []),
  ].filter((image) => isSupportedModelImage(image.dataUrl));
}

function collectImageContent(payload: CompareRequestPayload) {
  const images = collectImages(payload);
  const limitedImages = images.slice(0, MAX_MODEL_IMAGES);
  return {
    total: images.length,
    sent: limitedImages.length,
    content: limitedImages.map<AiContent>((image) => ({
      type: "input_image",
      image_url: image.dataUrl,
    })),
  };
}

async function callRealModel({
  payload,
  prompt,
  includeImages = true,
}: PromptRequest) {
  const apiUrl = process.env.CHATGPT_API_URL;
  const apiKey = process.env.CHATGPT_API_KEY;
  const model = process.env.CHATGPT_MODEL || "gpt-5.5";
  if (!apiUrl || !apiKey) return null;

  const imagePayload = includeImages
    ? collectImageContent(payload)
    : { total: 0, sent: 0, content: [] as AiContent[] };
  const imageNote = imagePayload.total
    ? `\n\n图片发送说明：浏览器端会压缩图片；服务端最多发送 ${MAX_MODEL_IMAGES} 张图片给模型。本次可发送 ${imagePayload.total} 张，实际发送 ${imagePayload.sent} 张，超出的页面/图片仅使用文字摘要。`
    : "";
  const content: AiContent[] = [{ type: "input_text", text: `${prompt}${imageNote}` }];
  content.push(...imagePayload.content);

  const response = await fetch(apiUrl, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Authorization: `Bearer ${apiKey}`,
    },
    body: JSON.stringify({
      model,
      input: [
        {
          type: "message",
          role: "user",
          content,
        },
      ],
      stream: false,
    }),
  });

  const rawText = await response.text();
  if (!response.ok) {
    throw new Error(`真实模型接口调用失败：HTTP ${response.status}`);
  }
  const contentType = response.headers.get("content-type") ?? "";
  if (contentType.includes("text/event-stream") || rawText.includes("data:")) {
    return parseSseText(rawText);
  }
  try {
    return extractOutputText(JSON.parse(rawText)) || rawText;
  } catch {
    return rawText;
  }
}

async function repairOutput(
  rawText: string,
  payload: CompareRequestPayload,
  fallbackRows?: Parameters<typeof mockCompareRepair>[2],
) {
  const repairPrompt = [
    "下面是一次信息核对 AI 输出，但不是可解析的目标 JSON。请只返回一个合法 JSON 对象。",
    'schema: {"text": string, "columns": string[], "rows": string[][]}',
    "columns 必须是：状态 / 检查对象 / 检查字段 / 待核对值 / 参考资料值 / 结论 / 置信度 / 差异说明 / 证据来源 / 建议处理。",
    "原始坏输出：",
    rawText.slice(0, 12000),
  ].join("\n\n");
  const repaired = await callRealModel({
    payload,
    prompt: repairPrompt,
    includeImages: false,
  });
  if (repaired) return parseCompareJson(repaired);
  return mockCompareRepair(rawText, payload, fallbackRows);
}

async function parseOrRepair(
  rawText: string,
  payload: CompareRequestPayload,
  mode: ModelMode,
  fallbackRows?: Parameters<typeof mockCompareRepair>[2],
): Promise<ParsedModelOutput> {
  try {
    return { table: parseCompareJson(rawText), repaired: false, mode };
  } catch {
    return {
      table: await repairOutput(rawText, payload, fallbackRows),
      repaired: true,
      mode,
    };
  }
}

async function runSpreadsheetCompare(
  payload: CompareRequestPayload,
): Promise<SpreadsheetCompareResult> {
  const subjectRows = collectSubjectSpreadsheetRows(payload);
  const batches = chunkCompareRows(subjectRows, COMPARE_BATCH_SIZE);
  const batchTables: ExtractedTable[] = [];
  const fallbackRowIds: string[] = [];
  const preFallbackMissingRowIds: string[] = [];
  const missingEvidenceRowIds: string[] = [];
  const referenceCorpus = collectReferenceCorpus(payload);
  let repaired = false;
  let repairCount = 0;
  let rerunCount = 0;
  let mode: ModelMode = "mock";

  for (const [batchIndex, batchRows] of batches.entries()) {
    const prompt = buildCompareBatchPrompt(
      payload,
      batchRows,
      batchIndex,
      batches.length,
    );
    const realText = await callRealModel({
      payload,
      prompt,
      includeImages: batchIndex === 0,
    });
    const batchMode: ModelMode = realText ? "real" : "mock";
    if (realText) mode = "real";
    const rawText = realText ?? JSON.stringify(mockCompareRows(payload, batchRows));
    let parsed = await parseOrRepair(rawText, payload, batchMode, batchRows);
    if (parsed.repaired) {
      repaired = true;
      repairCount += 1;
    }

    let coverage = ensureRowCoverage(parsed.table, batchRows);
    for (
      let rerunIndex = 0;
      rerunIndex < MAX_BATCH_RERUNS &&
      batchMode === "real" &&
      coverage.preFallbackMissingRowIds.length > 0;
      rerunIndex += 1
    ) {
      rerunCount += 1;
      const rerunPrompt = [
        prompt,
        "覆盖校验器发现你上一轮漏掉了以下 rowId。请只针对这些 rowId 重新输出合法 JSON；每个 rowId 必须出现在检查对象中，不允许省略。",
        JSON.stringify(coverage.preFallbackMissingRowIds, null, 2),
      ].join("\n\n");
      const rerunText = await callRealModel({
        payload,
        prompt: rerunPrompt,
        includeImages: false,
      });
      if (!rerunText) break;
      const rerunRows = batchRows.filter((row) =>
        coverage.preFallbackMissingRowIds.includes(row.rowId),
      );
      const rerunParsed = await parseOrRepair(
        rerunText,
        payload,
        "real",
        rerunRows,
      );
      if (rerunParsed.repaired) {
        repaired = true;
        repairCount += 1;
      }
      parsed = {
        table: mergeCompareTables([parsed.table, rerunParsed.table]),
        repaired: parsed.repaired || rerunParsed.repaired,
        mode: "real",
      };
      coverage = ensureRowCoverage(parsed.table, batchRows);
    }

    const evidenceCoverage = ensureReferenceEvidenceCoverage(
      coverage.table,
      batchRows,
      referenceCorpus,
    );

    batchTables.push(evidenceCoverage.table);
    fallbackRowIds.push(...coverage.fallbackRowIds);
    preFallbackMissingRowIds.push(...coverage.preFallbackMissingRowIds);
    missingEvidenceRowIds.push(...evidenceCoverage.missingEvidenceRowIds);
  }

  const table = mergeCompareTables(batchTables);
  const coverage = createCompareCoverageReport(table, subjectRows, {
    batchCount: batches.length,
    repairCount,
    rerunCount,
    fallbackRowIds,
    preFallbackMissingRowIds,
    missingEvidenceRowIds,
    batchSize: COMPARE_BATCH_SIZE,
  });
  return { table, mode, repaired, coverage };
}

async function runFreeformCompare(
  payload: CompareRequestPayload,
): Promise<ParsedModelOutput> {
  if (payload.forceRepairTest) {
    return parseOrRepair(
      "这是一个故意损坏的 mock 输出，用来测试核对 JSON 修复。",
      payload,
      "mock",
    );
  }
  const realText = await callRealModel({
    payload,
    prompt: buildComparePrompt(payload),
    includeImages: true,
  });
  const mode: ModelMode = realText ? "real" : "mock";
  const rawText = realText ?? JSON.stringify(mockCompare(payload));
  return parseOrRepair(rawText, payload, mode);
}

export async function POST(request: Request) {
  const cookieStore = await cookies();
  if (cookieStore.get("info_extractor_auth")?.value !== "ok") {
    return NextResponse.json({ error: "未登录" }, { status: 401 });
  }

  const payload = (await request.json()) as CompareRequestPayload;
  try {
    const subjectRows = collectSubjectSpreadsheetRows(payload);
    if (subjectRows.length && !payload.forceRepairTest) {
      const result = await runSpreadsheetCompare(payload);
      return NextResponse.json({
        ...result.table,
        mode: result.mode,
        repaired: result.repaired,
        coverage: result.coverage,
      });
    }

    const result = await runFreeformCompare(payload);
    return NextResponse.json({
      ...result.table,
      mode: result.mode,
      repaired: result.repaired,
    });
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "核对失败" },
      { status: 500 },
    );
  }
}
