import { cookies } from "next/headers";
import { NextResponse } from "next/server";
import {
  buildExtractionPrompt,
  mockExtract,
  mockRepair,
  parseExtractedJson,
} from "@/lib/extraction";
import type { ExtractRequestPayload } from "@/lib/types";

export const runtime = "nodejs";

const MAX_MODEL_IMAGES = 8;
const MIN_IMAGE_DATA_URL_LENGTH = 1_500;
const MAX_IMAGE_DATA_URL_LENGTH = 6_500_000;

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

function collectImageContent(payload: ExtractRequestPayload) {
  const images = [
    ...(payload.template.screenshotTemplate
      ? [payload.template.screenshotTemplate]
      : []),
    ...payload.files.flatMap((file) => file.images ?? []),
  ].filter((image) => isSupportedModelImage(image.dataUrl));
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

async function callRealModel(
  payload: ExtractRequestPayload,
  repairPrompt?: string,
) {
  const apiUrl = process.env.CHATGPT_API_URL;
  const apiKey = process.env.CHATGPT_API_KEY;
  const model = process.env.CHATGPT_MODEL || "gpt-5.5";
  if (!apiUrl || !apiKey) return null;

  const imagePayload = repairPrompt
    ? { total: 0, sent: 0, content: [] as AiContent[] }
    : collectImageContent(payload);
  const imageNote = imagePayload.total
    ? `\n\n图片发送说明：浏览器端会压缩图片；服务端最多发送 ${MAX_MODEL_IMAGES} 张图片给模型。本次可发送 ${imagePayload.total} 张，实际发送 ${imagePayload.sent} 张，超出的页面/图片仅使用文字摘要。`
    : "";
  const prompt = `${repairPrompt ?? buildExtractionPrompt(payload)}${imageNote}`;
  const content: AiContent[] = [{ type: "input_text", text: prompt }];
  if (!repairPrompt) content.push(...imagePayload.content);

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

async function repairOutput(rawText: string, payload: ExtractRequestPayload) {
  const repairPrompt = [
    "下面是一次 AI 输出，但不是可解析的目标 JSON。请只返回一个合法 JSON 对象。",
    'schema: {"text": string, "columns": string[], "rows": string[][]}',
    "原始坏输出：",
    rawText.slice(0, 12000),
  ].join("\n\n");
  const repaired = await callRealModel(payload, repairPrompt);
  if (repaired) return parseExtractedJson(repaired);
  return mockRepair(rawText, payload);
}

export async function POST(request: Request) {
  const cookieStore = await cookies();
  if (cookieStore.get("info_extractor_auth")?.value !== "ok") {
    return NextResponse.json({ error: "未登录" }, { status: 401 });
  }

  const payload = (await request.json()) as ExtractRequestPayload;
  let mode: "mock" | "real" = "mock";
  let repaired = false;
  try {
    let rawText: string;
    if (payload.forceRepairTest) {
      rawText = "这是一个故意损坏的 mock 输出，用来测试 JSON 修复。";
    } else {
      const realText = await callRealModel(payload);
      if (realText) {
        mode = "real";
        rawText = realText;
      } else {
        rawText = JSON.stringify(mockExtract(payload));
      }
    }

    try {
      const result = parseExtractedJson(rawText);
      return NextResponse.json({ ...result, mode, repaired });
    } catch {
      repaired = true;
      const result = await repairOutput(rawText, payload);
      return NextResponse.json({ ...result, mode, repaired });
    }
  } catch (error) {
    return NextResponse.json(
      { error: error instanceof Error ? error.message : "提取失败" },
      { status: 500 },
    );
  }
}
