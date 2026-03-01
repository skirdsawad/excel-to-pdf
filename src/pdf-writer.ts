import path from "node:path";
import type { Content, ContentTable, CustomTableLayout, TableCell } from "pdfmake/interfaces";
import { HorizontalAlignment, type CellData, type SheetData } from "./types.js";

// pdfmake exports a singleton instance via module.exports;
// named/destructured imports lose `this` context, so we import the whole object.
// eslint-disable-next-line @typescript-eslint/no-require-imports
const pdfmake = require("pdfmake") as typeof import("pdfmake");

const POINTS_PER_EXCEL_WIDTH = 5.25;
const A4_LANDSCAPE_WIDTH = 842;
const PAGE_MARGIN = 40;
const USABLE_WIDTH = A4_LANDSCAPE_WIDTH - PAGE_MARGIN * 2;
const DEFAULT_FONT_SIZE = 9;
const MIN_COL_WIDTH = 20;

function initFonts(): void {
  const pdfmakeRoot = path.dirname(
    require.resolve("pdfmake/package.json")
  );
  const fontDir = path.join(pdfmakeRoot, "fonts", "Roboto");

  pdfmake.setFonts({
    Roboto: {
      normal: path.join(fontDir, "Roboto-Regular.ttf"),
      bold: path.join(fontDir, "Roboto-Medium.ttf"),
      italics: path.join(fontDir, "Roboto-Italic.ttf"),
      bolditalics: path.join(fontDir, "Roboto-MediumItalic.ttf"),
    },
  });
}

function computeColumnWidths(excelWidths: number[]): number[] {
  const rawWidths = excelWidths.map(
    (w) => Math.max(w * POINTS_PER_EXCEL_WIDTH, MIN_COL_WIDTH)
  );
  const totalRaw = rawWidths.reduce((sum, w) => sum + w, 0);

  if (totalRaw <= 0) {
    return excelWidths.map(() => USABLE_WIDTH / excelWidths.length);
  }

  const scale = USABLE_WIDTH / totalRaw;

  return rawWidths.map((w) => w * scale);
}

function mapAlignment(
  alignment: HorizontalAlignment | undefined
): "left" | "center" | "right" | "justify" | undefined {
  if (!alignment) {
    return undefined;
  }

  switch (alignment) {
    case HorizontalAlignment.Left:
      return "left";
    case HorizontalAlignment.Center:
      return "center";
    case HorizontalAlignment.Right:
      return "right";
    case HorizontalAlignment.Justify:
      return "justify";
  }
}

function buildPdfCell(cell: CellData): TableCell {
  if (cell.isMergedSlave) {
    return {};
  }

  const pdfCell: Record<string, unknown> = {
    text: cell.value,
    fontSize: cell.style.fontSize ?? DEFAULT_FONT_SIZE,
  };

  if (cell.style.bold) {
    pdfCell.bold = true;
  }

  if (cell.style.italic) {
    pdfCell.italics = true;
  }

  if (cell.style.fontColor) {
    pdfCell.color = cell.style.fontColor;
  }

  if (cell.style.backgroundColor) {
    pdfCell.fillColor = cell.style.backgroundColor;
  }

  const alignment = mapAlignment(cell.style.horizontalAlignment);
  if (alignment) {
    pdfCell.alignment = alignment;
  }

  if (cell.style.underline) {
    pdfCell.decoration = "underline";
  }

  if (cell.style.strikethrough) {
    pdfCell.decoration = "lineThrough";
  }

  if (cell.colSpan > 1) {
    pdfCell.colSpan = cell.colSpan;
  }

  if (cell.rowSpan > 1) {
    pdfCell.rowSpan = cell.rowSpan;
  }

  // Build per-cell border array: [left, top, right, bottom]
  const borders = cell.style.borders;
  if (borders) {
    pdfCell.border = [
      !!borders.left,
      !!borders.top,
      !!borders.right,
      !!borders.bottom,
    ];

    const borderColor = [
      borders.left?.color ?? "#000000",
      borders.top?.color ?? "#000000",
      borders.right?.color ?? "#000000",
      borders.bottom?.color ?? "#000000",
    ];
    pdfCell.borderColor = borderColor;
  }

  if (cell.style.wrapText) {
    pdfCell.noWrap = false;
  }

  return pdfCell as unknown as TableCell;
}

function buildSheetTable(sheet: SheetData): ContentTable {
  const widths = computeColumnWidths(sheet.columnWidths);

  const body: TableCell[][] = sheet.rows.map((row) =>
    row.map((cell) => buildPdfCell(cell))
  );

  // Ensure the table has at least one row
  if (body.length === 0) {
    body.push(
      Array.from({ length: sheet.totalColumns }, () => ({ text: "" }))
    );
  }

  const tableLayout: CustomTableLayout = {
    hLineWidth(rowIndex: number): number {
      return 0.5;
    },
    vLineWidth(columnIndex: number): number {
      return 0.5;
    },
    hLineColor(): string {
      return "#cccccc";
    },
    vLineColor(): string {
      return "#cccccc";
    },
    paddingLeft(): number {
      return 2;
    },
    paddingRight(): number {
      return 2;
    },
    paddingTop(): number {
      return 1;
    },
    paddingBottom(): number {
      return 1;
    },
  };

  return {
    table: {
      headerRows: 0,
      widths,
      body,
      dontBreakRows: true,
    },
    layout: tableLayout,
  };
}

export async function generatePdf(
  sheets: SheetData[],
  outputPath: string
): Promise<void> {
  initFonts();

  const content: Content[] = [];

  for (let i = 0; i < sheets.length; i++) {
    const sheet = sheets[i];

    // Add sheet name as a header
    content.push({
      text: sheet.name,
      fontSize: 12,
      bold: true,
      margin: [0, 0, 0, 6] as [number, number, number, number],
    });

    // Add the table
    content.push(buildSheetTable(sheet));

    // Page break after each sheet except the last
    if (i < sheets.length - 1) {
      content.push({
        text: "",
        pageBreak: "after",
      });
    }
  }

  const docDefinition = {
    pageSize: "A4" as const,
    pageOrientation: "landscape" as const,
    pageMargins: [PAGE_MARGIN, PAGE_MARGIN, PAGE_MARGIN, PAGE_MARGIN] as [
      number,
      number,
      number,
      number,
    ],
    content,
    defaultStyle: {
      font: "Roboto",
      fontSize: DEFAULT_FONT_SIZE,
    },
  };

  const pdf = pdfmake.createPdf(docDefinition);
  await pdf.write(outputPath);
}
