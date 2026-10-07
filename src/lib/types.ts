export type ExtractedTable = {
  text: string;
  columns: string[];
  rows: string[][];
};

export type SourceImage = {
  name: string;
  mimeType: string;
  dataUrl: string;
};

export type SourceTable = {
  sheetName: string;
  columns: string[];
  rows: string[][];
};

export type SourceFilePayload = {
  kind: "image" | "pdf" | "excel";
  name: string;
  text?: string;
  pageCount?: number;
  images?: SourceImage[];
  tables?: SourceTable[];
};

export type TemplatePayload = {
  instruction: string;
  excelTemplate?: SourceTable;
  screenshotTemplate?: SourceImage;
  outputMode: "text" | "excel" | "both";
};

export type ExtractRequestPayload = {
  sourceText: string;
  files: SourceFilePayload[];
  template: TemplatePayload;
  forceRepairTest?: boolean;
};

export type ExtractResponsePayload = ExtractedTable & {
  mode: "mock" | "real";
  repaired: boolean;
};

export type CompareSidePayload = {
  text: string;
  files: SourceFilePayload[];
};

export type CompareRequestPayload = {
  subject: CompareSidePayload;
  reference: CompareSidePayload;
  requirement: string;
  templateName?: string;
  forceRepairTest?: boolean;
};

export type CompareCoveragePayload = {
  inputRowCount: number;
  expectedRowIds: string[];
  outputRowIds: string[];
  coveredRowIds: string[];
  missingRowIds: string[];
  missingRowCount: number;
  preFallbackMissingRowIds: string[];
  preFallbackMissingRowCount: number;
  fallbackRowIds: string[];
  fallbackCount: number;
  fallbackStatus: "not_needed" | "filled_missing_rows";
  missingEvidenceRowIds: string[];
  missingEvidenceRowCount: number;
  batchCount: number;
  batchSize: number;
  repairCount: number;
  rerunCount: number;
};

export type CompareResponsePayload = ExtractedTable & {
  mode: "mock" | "real";
  repaired: boolean;
  coverage?: CompareCoveragePayload;
};
