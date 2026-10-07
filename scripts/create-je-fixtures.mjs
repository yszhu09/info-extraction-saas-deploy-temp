import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import * as XLSX from "@e965/xlsx";
import ExcelJS from "exceljs";
import { PDFDocument, StandardFonts, rgb } from "pdf-lib";

async function writeWorkbook(filePath, rows) {
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("Sheet1");
  rows.forEach((row) => sheet.addRow(row));
  sheet.columns.forEach((column) => {
    column.width = Math.max(
      12,
      ...column.values.map((value) => String(value ?? "").length),
    );
  });
  await workbook.xlsx.writeFile(filePath);
}

async function writeLegacyWorkbook(filePath, rows) {
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(
    workbook,
    XLSX.utils.aoa_to_sheet(rows),
    "Sheet1",
  );
  return fs.writeFile(
    filePath,
    XLSX.write(workbook, { type: "buffer", bookType: "biff8" }),
  );
}

async function writeCsv(filePath, rows) {
  const escapeCell = (value) => {
    const text = String(value ?? "");
    return /[",\n]/.test(text) ? `"${text.replaceAll('"', '""')}"` : text;
  };
  await fs.writeFile(
    filePath,
    rows.map((row) => row.map(escapeCell).join(",")).join("\n"),
  );
}

async function writePdf(filePath, title, lines) {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([595, 842]);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  page.drawText(title, {
    x: 56,
    y: 780,
    size: 18,
    font,
    color: rgb(0.03, 0.11, 0.22),
  });
  lines.forEach((line, index) => {
    page.drawText(line, {
      x: 56,
      y: 736 - index * 28,
      size: 13,
      font,
      color: rgb(0.05, 0.12, 0.22),
    });
  });
  await fs.writeFile(filePath, await pdf.save());
}

export async function createJeFixtures(
  outDir = path.resolve(process.env.OUT_DIR || "validation-output/task66"),
) {
  await fs.mkdir(outDir, { recursive: true });

  const phase1SourceRows = [
    ["Item", "Version", "Quantity", "Note"],
    ["A-001", "Rev.B", "2", "source table row"],
    ["A-002", "Rev.A", "4", "source table row"],
    ["A-004", "Unknown", "1", "requires confirmation"],
  ];
  await writeWorkbook(path.join(outDir, "phase1-source.xlsx"), phase1SourceRows);
  await writeLegacyWorkbook(
    path.join(outDir, "phase1-source-legacy.xls"),
    phase1SourceRows,
  );
  await writeCsv(path.join(outDir, "phase1-source-csv.csv"), phase1SourceRows);

  const phase1TemplateRows = [
    ["Part No", "Version", "Quantity", "Risk", "Evidence"],
    ["A-001", "", "", "", ""],
  ];
  await writeWorkbook(
    path.join(outDir, "phase1-template.xlsx"),
    phase1TemplateRows,
  );
  await writeLegacyWorkbook(
    path.join(outDir, "phase1-template-legacy.xls"),
    phase1TemplateRows,
  );

  const subjectRows = [
    ["Part No", "Version", "Quantity", "Material"],
    ["A-001", "Rev.B", "2", "AL6061"],
    ["A-002", "Rev.A", "4", "Steel"],
    ["A-003", "Rev.D", "1", "Copper"],
    ["A-004", "Rev.A", "6", "POM"],
    ["A-005", "Rev.C", "3", "Brass"],
    ["A-006", "Rev.A", "8", "Nylon"],
    ["A-007", "Rev.B", "2", "Titanium"],
    ["A-008", "Rev.A", "5", "Rubber"],
    ["A-009", "Rev.C", "7", "Glass"],
    ["A-010", "Rev.B", "9", "Ceramic"],
    ["A-011", "Rev.A", "1", "Carbon"],
    ["A-012", "Rev.D", "4", "PVC"],
  ];
  await writeWorkbook(path.join(outDir, "subject-bom.xlsx"), subjectRows);
  await writeLegacyWorkbook(path.join(outDir, "subject-bom-legacy.xls"), subjectRows);
  await writeCsv(path.join(outDir, "subject-bom.csv"), subjectRows);

  const referenceRows = [
    ["Part No", "Version", "Quantity", "Source"],
    ["A-001", "Rev.C", "2", "Drawing page 1"],
    ["A-002", "Rev.A", "4", "Released register"],
    ["A-003", "Rev.D", "", "quantity missing"],
    ["A-005", "Rev.C", "3", "Released register"],
    ["A-008", "Rev.A", "5", "Released register"],
  ];
  await writeWorkbook(path.join(outDir, "reference-register.xlsx"), referenceRows);
  await writeCsv(path.join(outDir, "reference-register.csv"), referenceRows);

  await writePdf(path.join(outDir, "phase1-source.pdf"), "Phase 1 Source PDF", [
    "Part A-001 version Rev.B quantity 2 material AL6061.",
    "Part A-002 version Rev.A quantity 4 material Steel.",
    "Part A-004 has uncertain version and needs manual confirmation.",
  ]);

  await writePdf(
    path.join(outDir, "reference-drawing.pdf"),
    "Reference Drawing PDF",
    [
      "Part A-001 version Rev.C quantity 2 material AL6061.",
      "Part A-002 version Rev.A quantity 4 material Steel.",
      "Part A-003 exists but quantity is not listed in released drawing.",
      "Part A-005 version Rev.C quantity 3 material Brass is released.",
      "Part A-008 version Rev.A quantity 5 material Rubber is released.",
      "If source is unclear, mark needs manual confirmation.",
    ],
  );

  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="420" height="220"><rect width="100%" height="100%" fill="white"/><text x="24" y="58" font-size="26" fill="#0f172a">JE smoke note</text><text x="24" y="108" font-size="24" fill="#0f172a">Part A-004</text><text x="24" y="154" font-size="20" fill="#64748b">version uncertain / manual check</text></svg>`;
  await fs.writeFile(path.join(outDir, "sample-note.svg"), svg);

  return {
    outDir,
    phase1SourcePdf: path.join(outDir, "phase1-source.pdf"),
    phase1SourceXlsx: path.join(outDir, "phase1-source.xlsx"),
    phase1SourceLegacyXls: path.join(outDir, "phase1-source-legacy.xls"),
    phase1SourceCsv: path.join(outDir, "phase1-source-csv.csv"),
    phase1TemplateXlsx: path.join(outDir, "phase1-template.xlsx"),
    phase1TemplateLegacyXls: path.join(outDir, "phase1-template-legacy.xls"),
    subjectXlsx: path.join(outDir, "subject-bom.xlsx"),
    subjectLegacyXls: path.join(outDir, "subject-bom-legacy.xls"),
    subjectCsv: path.join(outDir, "subject-bom.csv"),
    referenceXlsx: path.join(outDir, "reference-register.xlsx"),
    referenceCsv: path.join(outDir, "reference-register.csv"),
    referencePdf: path.join(outDir, "reference-drawing.pdf"),
    sampleSvg: path.join(outDir, "sample-note.svg"),
  };
}

const invokedPath = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (invokedPath && fileURLToPath(import.meta.url) === invokedPath) {
  const result = await createJeFixtures();
  console.log(JSON.stringify(result, null, 2));
}
