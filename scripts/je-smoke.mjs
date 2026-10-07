import fs from "node:fs/promises";
import path from "node:path";
import ExcelJS from "exceljs";
import { createJeFixtures } from "./create-je-fixtures.mjs";

const baseUrl = process.env.BASE_URL || "http://127.0.0.1:3015";
const outDir = path.resolve(process.env.OUT_DIR || "validation-output/task66");
const expectRealAi = process.env.EXPECT_REAL_AI === "1";
const fullUnderBlockedOldChunks = process.env.BLOCK_OLD_NEXT_CHUNKS !== "0";
const oldChunkPrefix = "/_next/static/chunks/";
const safeStaticPrefix = "/infowb-assets/_next/static/";

async function loadPlaywright() {
  try {
    const playwrightModule = await import("playwright");
    return playwrightModule.chromium ? playwrightModule : playwrightModule.default;
  } catch {
    const playwrightModule = await import(
      "/root/.nvm/versions/node/v22.22.0/lib/node_modules/playwright/index.js"
    );
    return playwrightModule.chromium ? playwrightModule : playwrightModule.default;
  }
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

async function parseWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);
  const sheet = workbook.worksheets[0];
  const rows = [];
  sheet.eachRow((row) => {
    rows.push(row.values.slice(1).map((value) => String(value ?? "")));
  });
  return {
    sheetName: sheet.name,
    rowCount: rows.length,
    rows,
    previewRows: rows.slice(0, 8),
  };
}

function expectedCompareRowIds(fixtures) {
  return [fixtures.subjectXlsx, fixtures.subjectLegacyXls, fixtures.subjectCsv]
    .flatMap((filePath) =>
      Array.from({ length: 12 }, (_, index) =>
        `${path.basename(filePath)}#Sheet1!R${index + 2}`,
      ),
    );
}

function collectWorkbookRowIds(workbook) {
  const text = workbook.rows.flat().join("\n");
  return Array.from(
    new Set(
      text.match(/subject-bom(?:-legacy)?\.(?:xlsx|xls|csv)#Sheet1!R\d+/g) ?? [],
    ),
  );
}

function expectedMissingEvidenceRowIds(fixtures) {
  const missingExcelRows = [5, 7, 8, 10, 11, 12, 13];
  return [fixtures.subjectXlsx, fixtures.subjectLegacyXls, fixtures.subjectCsv]
    .flatMap((filePath) =>
      missingExcelRows.map(
        (rowNumber) => `${path.basename(filePath)}#Sheet1!R${rowNumber}`,
      ),
    );
}

function collectManualOrMissingRowIds(workbook) {
  return workbook.rows
    .slice(1)
    .filter((row) => ["需人工确认", "右侧缺失"].includes(row[5]))
    .flatMap(
      (row) =>
        row
          .join("\n")
          .match(/subject-bom(?:-legacy)?\.(?:xlsx|xls|csv)#Sheet1!R\d+/g) ?? [],
    );
}

function assertPartialReferenceStatuses({ expectedRowIds, workbook }) {
  const actualSet = new Set(collectManualOrMissingRowIds(workbook));
  const missing = expectedRowIds.filter((rowId) => !actualSet.has(rowId));
  assert(
    missing.length === 0,
    `partial-reference rows not marked manual/missing: ${missing.join(", ")}`,
  );
  return {
    checkedRowIds: expectedRowIds,
    rowIds: Array.from(actualSet),
    missingRowIds: missing,
  };
}

function assertRowIdCoverage({ expectedRowIds, actualRowIds, label }) {
  const actualSet = new Set(actualRowIds);
  const missing = expectedRowIds.filter((rowId) => !actualSet.has(rowId));
  assert(
    missing.length === 0,
    `${label} missing rowIds: ${missing.join(", ")}`,
  );
  return { actualRowIds, missingRowIds: missing };
}

async function addAuthCookie(context) {
  const url = new URL(baseUrl);
  await context.addCookies([
    {
      name: "info_extractor_auth",
      value: "ok",
      domain: url.hostname,
      path: "/",
      httpOnly: true,
      sameSite: "Lax",
      secure: url.protocol === "https:",
      expires: Math.floor(Date.now() / 1000) + 60 * 60,
    },
  ]);
}

async function waitForReady(page) {
  await page.getByTestId("infowb-inline-probe").waitFor({ timeout: 30000 });
  await page
    .getByTestId("infowb-je-status")
    .filter({ hasText: "浏览器功能已就绪" })
    .waitFor({ timeout: 30000 });
}

async function collectStaticResources(page) {
  return page.evaluate(() =>
    [
      ...Array.from(document.scripts).map((script) => script.src),
      ...Array.from(document.querySelectorAll("link[href]")).map(
        (link) => link.href,
      ),
    ].filter((url) => url.includes("/_next/static/")),
  );
}

function analyzeResources(urls) {
  const nextStatic = urls.map((url) => new URL(url));
  const unsafe = nextStatic.filter(
    (url) => !url.pathname.startsWith(safeStaticPrefix),
  );
  const chunks = nextStatic.filter((url) => url.pathname.includes("/chunks/"));
  const badChunkNames = chunks
    .map((url) => path.posix.basename(url.pathname))
    .filter((name) => /turbopack|~|\.\./i.test(name));
  return {
    total: nextStatic.length,
    unsafe: unsafe.map((url) => url.pathname),
    chunks: chunks.map((url) => url.pathname),
    badChunkNames,
  };
}

async function pasteImage(page, testId, filePath, name) {
  const svg = await fs.readFile(filePath, "utf8");
  const dataUrl = `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`;
  await page.getByTestId(testId).focus();
  await page.evaluate(
    async ({ testId: targetTestId, imageDataUrl, fileName }) => {
      const element = document.querySelector(`[data-testid="${targetTestId}"]`);
      if (!element) throw new Error(`missing ${targetTestId}`);
      const blob = await (await fetch(imageDataUrl)).blob();
      const transfer = new DataTransfer();
      transfer.items.add(new File([blob], fileName, { type: "image/svg+xml" }));
      element.dispatchEvent(
        new ClipboardEvent("paste", {
          bubbles: true,
          cancelable: true,
          clipboardData: transfer,
        }),
      );
    },
    { testId, imageDataUrl: dataUrl, fileName: name },
  );
  await page.getByText(name).first().waitFor({ timeout: 30000 });
}

async function dropImage(page, testId, filePath, name) {
  const svg = await fs.readFile(filePath, "utf8");
  const dataUrl = `data:image/svg+xml;base64,${Buffer.from(svg).toString("base64")}`;
  await page.evaluate(
    async ({ testId: targetTestId, imageDataUrl, fileName }) => {
      const element = document.querySelector(`[data-testid="${targetTestId}"]`);
      if (!element) throw new Error(`missing ${targetTestId}`);
      const blob = await (await fetch(imageDataUrl)).blob();
      const transfer = new DataTransfer();
      transfer.items.add(new File([blob], fileName, { type: "image/svg+xml" }));
      for (const type of ["dragenter", "dragover", "drop"]) {
        element.dispatchEvent(
          new DragEvent(type, {
            bubbles: true,
            cancelable: true,
            dataTransfer: transfer,
          }),
        );
      }
    },
    { testId, imageDataUrl: dataUrl, fileName: name },
  );
  await page.getByText(name).first().waitFor({ timeout: 30000 });
}

async function clickAndParseDownload(page, buttonName, filePath) {
  const [download] = await Promise.all([
    page.waitForEvent("download", { timeout: 60000 }),
    page.getByRole("button", { name: buttonName }).click(),
  ]);
  await download.saveAs(filePath);
  return {
    suggestedFilename: download.suggestedFilename(),
    workbook: await parseWorkbook(filePath),
  };
}

async function runScenario({ name, blockOldChunks, fullFlow, fixtures }) {
  const { chromium } = await loadPlaywright();
  const browser = await chromium.launch({ headless: true });
  const context = await browser.newContext({ acceptDownloads: true });
  await context.grantPermissions(["clipboard-read", "clipboard-write"], {
    origin: new URL(baseUrl).origin,
  });
  await addAuthCookie(context);

  const summary = {
    name,
    blockOldChunks,
    fullFlow,
    consoleErrors: [],
    pageErrors: [],
    failedRequests: [],
    badResponses: [],
    blockedOldChunkRequests: [],
    pages: {},
  };

  if (blockOldChunks) {
    await context.route("**/*", async (route) => {
      const requestUrl = new URL(route.request().url());
      if (requestUrl.pathname.startsWith(oldChunkPrefix)) {
        summary.blockedOldChunkRequests.push(requestUrl.pathname);
        await route.fulfill({
          status: 403,
          contentType: "text/plain",
          body: "old Next.js chunk path blocked by JE smoke",
        });
        return;
      }
      await route.continue();
    });
  }

  const page = await context.newPage();
  page.setDefaultTimeout(30000);
  page.on("console", (message) => {
    if (message.type() === "error") summary.consoleErrors.push(message.text());
  });
  page.on("pageerror", (error) => summary.pageErrors.push(error.message));
  page.on("requestfailed", (request) => {
    summary.failedRequests.push({
      url: request.url(),
      failure: request.failure()?.errorText ?? "unknown",
    });
  });
  page.on("response", (response) => {
    if (response.status() >= 400) {
      summary.badResponses.push({ status: response.status(), url: response.url() });
    }
  });

  try {
    await page.goto(`${baseUrl}/`, { waitUntil: "domcontentloaded" });
    await waitForReady(page);
    const rootResources = await collectStaticResources(page);
    const rootResourceAnalysis = analyzeResources(rootResources);
    assert(rootResourceAnalysis.total > 0, "root has no Next static resources");
    assert(
      rootResourceAnalysis.unsafe.length === 0,
      `root has unsafe static URLs: ${rootResourceAnalysis.unsafe.join(", ")}`,
    );
    assert(
      rootResourceAnalysis.badChunkNames.length === 0,
      `root has Turbopack-style chunks: ${rootResourceAnalysis.badChunkNames.join(", ")}`,
    );
    summary.pages.root = {
      url: page.url(),
      titleVisible: await page.getByText("信息提取工作台").isVisible(),
      probeText: await page.getByTestId("infowb-inline-probe").textContent(),
      readyText: await page.getByTestId("infowb-je-status").textContent(),
      resources: rootResourceAnalysis,
    };

    if (fullFlow) {
      await page
        .getByTestId("extract-source-textarea")
        .fill(
          "Phase 1 text: A-001 Rev.B quantity 2; A-002 Rev.A quantity 4; A-004 version uncertain and requires manual confirmation.",
        );
      await pasteImage(
        page,
        "extract-source-textarea",
        fixtures.sampleSvg,
        "phase1-source-pasted.svg",
      );
      await dropImage(
        page,
        "extract-source-dropzone",
        fixtures.sampleSvg,
        "phase1-source-dropped.svg",
      );
      await page
        .getByTestId("extract-source-file-input")
        .setInputFiles([
          fixtures.phase1SourcePdf,
          fixtures.sampleSvg,
          fixtures.phase1SourceXlsx,
          fixtures.phase1SourceLegacyXls,
          fixtures.phase1SourceCsv,
        ]);
      await page.getByText("phase1-source.pdf").waitFor({ timeout: 60000 });
      await page.getByText("sample-note.svg").first().waitFor({ timeout: 60000 });
      await page.getByText("phase1-source.xlsx").waitFor({ timeout: 60000 });
      await page.getByText("phase1-source-legacy.xls").waitFor({ timeout: 60000 });
      await page.getByText("phase1-source-csv.csv").waitFor({ timeout: 60000 });
      await page
        .getByTestId("extract-template-textarea")
        .fill("请提取 Part No、Version、Quantity、Risk、Evidence，低置信度标记需人工确认。");
      await pasteImage(
        page,
        "extract-template-textarea",
        fixtures.sampleSvg,
        "phase1-template-pasted.svg",
      );
      await page
        .getByTestId("extract-template-file-input")
        .setInputFiles(fixtures.phase1TemplateLegacyXls);
      await page
        .getByText("Excel / CSV 模板：Part No / Version / Quantity / Risk / Evidence", { exact: true })
        .first()
        .waitFor({ timeout: 60000 });

      const [extractResponse] = await Promise.all([
        page.waitForResponse(
          (response) =>
            response.url().includes("/api/extract") &&
            response.request().method() === "POST",
          { timeout: 180000 },
        ),
        page.getByTestId("extract-start").click(),
      ]);
      const extractBody = await extractResponse.json();
      assert(extractResponse.ok(), extractBody.error || "extract api failed");
      if (expectRealAi) assert(extractBody.mode === "real", "extract did not use real AI");
      await page.getByTestId("extract-result-table").waitFor({ timeout: 180000 });
      const extractText = await page.getByTestId("extract-result-table").innerText();
      assert(extractText.length > 20, "extract table is too short");
      await page.getByRole("button", { name: "复制表格" }).first().click();
      const extractClipboard = await page.evaluate(() => navigator.clipboard.readText());
      assert(extractClipboard.length > 20, "extract copy result is too short");
      const extractDownload = await clickAndParseDownload(
        page,
        "下载 .xlsx",
        path.join(outDir, `${name}-phase1-extract.xlsx`),
      );
      assert(extractDownload.workbook.rowCount >= 2, "extract xlsx has too few rows");
      summary.pages.root.extract = {
        mode: extractBody.mode,
        repaired: extractBody.repaired,
        tableLength: extractText.length,
        copyLength: extractClipboard.length,
        download: extractDownload,
      };
    }

    await page.goto(`${baseUrl}/compare`, { waitUntil: "domcontentloaded" });
    await waitForReady(page);
    const compareResources = await collectStaticResources(page);
    const compareResourceAnalysis = analyzeResources(compareResources);
    assert(compareResourceAnalysis.total > 0, "compare has no Next static resources");
    assert(
      compareResourceAnalysis.unsafe.length === 0,
      `compare has unsafe static URLs: ${compareResourceAnalysis.unsafe.join(", ")}`,
    );
    assert(
      compareResourceAnalysis.badChunkNames.length === 0,
      `compare has Turbopack-style chunks: ${compareResourceAnalysis.badChunkNames.join(", ")}`,
    );
    summary.pages.compare = {
      url: page.url(),
      titleVisible: await page.getByText("信息核对工作台").isVisible(),
      probeText: await page.getByTestId("infowb-inline-probe").textContent(),
      readyText: await page.getByTestId("infowb-je-status").textContent(),
      resources: compareResourceAnalysis,
    };

    if (fullFlow) {
      await page
        .getByTestId("subject-textarea")
        .fill("Subject text: A-001 Rev.B quantity 2; A-002 Rev.A quantity 4; A-003 Rev.D quantity 1; A-004 through A-012 must still each be checked even when reference evidence is partial.");
      await page
        .getByTestId("reference-textarea")
        .fill("Reference text: A-001 Rev.C quantity 2; A-002 Rev.A quantity 4; A-003 quantity missing in drawing.");
      await pasteImage(page, "subject-textarea", fixtures.sampleSvg, "subject-pasted.svg");
      await dropImage(page, "reference-dropzone", fixtures.sampleSvg, "reference-dropped.svg");
      await page
        .getByTestId("subject-file-input")
        .setInputFiles([
          fixtures.subjectXlsx,
          fixtures.subjectLegacyXls,
          fixtures.subjectCsv,
          fixtures.sampleSvg,
        ]);
      await page.getByText("subject-bom.xlsx").waitFor({ timeout: 60000 });
      await page.getByText("subject-bom-legacy.xls").waitFor({ timeout: 60000 });
      await page.getByText("subject-bom.csv").waitFor({ timeout: 60000 });
      await page.getByText("sample-note.svg").first().waitFor({ timeout: 60000 });
      await page
        .getByTestId("reference-file-input")
        .setInputFiles([
          fixtures.referencePdf,
          fixtures.referenceXlsx,
          fixtures.referenceCsv,
        ]);
      await page.getByText("reference-drawing.pdf").waitFor({ timeout: 60000 });
      await page.getByText("reference-register.xlsx").waitFor({ timeout: 60000 });
      await page.getByText("reference-register.csv").waitFor({ timeout: 60000 });
      await page
        .getByTestId("compare-requirement")
        .fill("核对零件号、版本、数量、材料/依据，缺失或依据不明确必须标记需人工确认。");

      const [compareResponse] = await Promise.all([
        page.waitForResponse(
          (response) =>
            response.url().includes("/api/compare") &&
            response.request().method() === "POST",
          { timeout: 180000 },
        ),
        page.getByTestId("compare-start").click(),
      ]);
      const compareBody = await compareResponse.json();
      assert(compareResponse.ok(), compareBody.error || "compare api failed");
      if (expectRealAi) assert(compareBody.mode === "real", "compare did not use real AI");
      const expectedRowIds = expectedCompareRowIds(fixtures);
      assert(
        compareBody.coverage?.inputRowCount === expectedRowIds.length,
        `compare input coverage count mismatch: ${compareBody.coverage?.inputRowCount}`,
      );
      assert(
        compareBody.coverage?.missingRowCount === 0,
        `compare API missing rowIds: ${(compareBody.coverage?.missingRowIds ?? []).join(", ")}`,
      );
      assert(
        compareBody.coverage?.expectedRowIds?.length === expectedRowIds.length,
        "compare API expected rowId list length mismatch",
      );
      assertRowIdCoverage({
        expectedRowIds,
        actualRowIds: compareBody.coverage?.outputRowIds ?? [],
        label: "compare API coverage",
      });
      const missingEvidenceRowIds = expectedMissingEvidenceRowIds(fixtures);
      assertRowIdCoverage({
        expectedRowIds: missingEvidenceRowIds,
        actualRowIds: compareBody.coverage?.missingEvidenceRowIds ?? [],
        label: "compare API missing-evidence coverage",
      });
      await page.getByTestId("compare-result-table").waitFor({ timeout: 180000 });
      const compareText = await page.getByTestId("compare-result-table").innerText();
      assert(compareText.length > 20, "compare table is too short");
      assertRowIdCoverage({
        expectedRowIds,
        actualRowIds: expectedRowIds.filter((rowId) => compareText.includes(rowId)),
        label: "compare page table",
      });
      await page.getByRole("button", { name: "复制表格" }).last().click();
      const compareClipboard = await page.evaluate(() => navigator.clipboard.readText());
      assert(compareClipboard.length > 20, "compare copy result is too short");
      const compareDownload = await clickAndParseDownload(
        page,
        "下载 .xlsx",
        path.join(outDir, `${name}-phase2-compare.xlsx`),
      );
      assert(
        compareDownload.workbook.rowCount >= expectedRowIds.length + 1,
        "compare xlsx has fewer rows than full subject coverage requires",
      );
      const workbookCoverage = assertRowIdCoverage({
        expectedRowIds,
        actualRowIds: collectWorkbookRowIds(compareDownload.workbook),
        label: "downloaded comparison-check.xlsx",
      });
      const partialReferenceStatusCoverage = assertPartialReferenceStatuses({
        expectedRowIds: missingEvidenceRowIds,
        workbook: compareDownload.workbook,
      });
      summary.pages.compare.compare = {
        mode: compareBody.mode,
        repaired: compareBody.repaired,
        coverage: compareBody.coverage,
        expectedRowIds,
        downloadedWorkbookCoverage: workbookCoverage,
        partialReferenceStatusCoverage,
        tableLength: compareText.length,
        copyLength: compareClipboard.length,
        download: compareDownload,
      };
    }

    assert(summary.consoleErrors.length === 0, `console errors: ${summary.consoleErrors.join(" | ")}`);
    assert(summary.pageErrors.length === 0, `page errors: ${summary.pageErrors.join(" | ")}`);
    assert(summary.failedRequests.length === 0, `failed requests: ${JSON.stringify(summary.failedRequests)}`);
    assert(summary.badResponses.length === 0, `bad responses: ${JSON.stringify(summary.badResponses)}`);
    assert(
      !blockOldChunks || summary.blockedOldChunkRequests.length === 0,
      `app still requested blocked old chunks: ${summary.blockedOldChunkRequests.join(", ")}`,
    );
    return summary;
  } finally {
    await browser.close();
  }
}

await fs.mkdir(outDir, { recursive: true });
const fixtures = await createJeFixtures(outDir);
const scenarios = [
  { name: "normal", blockOldChunks: false, fullFlow: false, fixtures },
  {
    name: "old-next-blocked",
    blockOldChunks: fullUnderBlockedOldChunks,
    fullFlow: true,
    fixtures,
  },
];
const results = [];
for (const scenario of scenarios) {
  results.push(await runScenario(scenario));
}
const result = {
  baseUrl,
  expectRealAi,
  safeStaticPrefix,
  oldChunkPrefix,
  results,
};
const summaryPath = path.join(outDir, process.env.SUMMARY_FILE || "je-smoke-summary.json");
await fs.writeFile(summaryPath, JSON.stringify(result, null, 2));
console.log(JSON.stringify(result, null, 2));
