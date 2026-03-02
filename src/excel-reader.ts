import ExcelJS from "exceljs";
import { mapExcelStyle } from "./style-mapper.js";
import type { CellData, MergeInfo, SheetData } from "./types.js";

const DEFAULT_COLUMN_WIDTH = 8.43;
const DEFAULT_ROW_HEIGHT = 15;

function colLetterToNumber(letters: string): number {
  let result = 0;
  for (const char of letters.toUpperCase()) {
    result = result * 26 + (char.charCodeAt(0) - "A".charCodeAt(0) + 1);
  }

  return result;
}

interface CellRef {
  row: number;
  col: number;
}

function parseCellRef(ref: string): CellRef {
  const match = ref.match(/^\$?([A-Z]+)\$?(\d+)$/);
  if (!match) {
    throw new Error(`Invalid cell reference: ${ref}`);
  }

  return {
    col: colLetterToNumber(match[1]),
    row: parseInt(match[2], 10),
  };
}

interface MergeRange {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}

function parseMergeRange(range: string): MergeRange {
  const parts = range.split(":");
  if (parts.length !== 2) {
    throw new Error(`Invalid merge range: ${range}`);
  }

  const start = parseCellRef(parts[0]);
  const end = parseCellRef(parts[1]);

  return {
    startRow: start.row,
    startCol: start.col,
    endRow: end.row,
    endCol: end.col,
  };
}

function buildMergeMap(
  merges: string[]
): Map<string, MergeInfo> {
  const map = new Map<string, MergeInfo>();

  for (const range of merges) {
    const { startRow, startCol, endRow, endCol } = parseMergeRange(range);
    const rowSpan = endRow - startRow + 1;
    const colSpan = endCol - startCol + 1;

    for (let r = startRow; r <= endRow; r++) {
      for (let c = startCol; c <= endCol; c++) {
        map.set(`${r},${c}`, {
          masterRow: startRow,
          masterCol: startCol,
          rowSpan,
          colSpan,
        });
      }
    }
  }

  return map;
}

function formatFormulaResult(result: unknown): string {
  if (result === undefined) {
    return "";
  }
  if (result instanceof Date) {
    return result.toLocaleDateString();
  }

  return String(result);
}

function extractCellValue(cell: ExcelJS.Cell): string {
  const value = cell.value;

  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "string") {
    return value;
  }

  if (typeof value === "number") {
    return String(value);
  }

  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }

  if (value instanceof Date) {
    return value.toLocaleDateString();
  }

  // Handle RichText
  if (typeof value === "object" && "richText" in value) {
    const richText = value as ExcelJS.CellRichTextValue;

    return richText.richText.map((part) => part.text).join("");
  }

  // Handle formula results
  if (typeof value === "object" && "formula" in value) {
    return formatFormulaResult((value as ExcelJS.CellFormulaValue).result);
  }

  // Handle shared formula
  if (typeof value === "object" && "sharedFormula" in value) {
    return formatFormulaResult((value as ExcelJS.CellSharedFormulaValue).result);
  }

  // Handle hyperlink
  if (typeof value === "object" && "hyperlink" in value) {
    const hyperlink = value as ExcelJS.CellHyperlinkValue;

    return hyperlink.text?.toString() ?? hyperlink.hyperlink ?? "";
  }

  // Handle error
  if (typeof value === "object" && "error" in value) {
    const error = value as ExcelJS.CellErrorValue;

    return error.error?.toString() ?? "#ERROR";
  }

  return String(value);
}

function readWorksheet(
  worksheet: ExcelJS.Worksheet,
  mergeMap: Map<string, MergeInfo>
): SheetData {
  const totalRows = worksheet.rowCount;
  const totalColumns = worksheet.columnCount;

  if (totalRows === 0 || totalColumns === 0) {
    return {
      name: worksheet.name,
      rows: [],
      columnWidths: [],
      rowHeights: [],
      totalColumns: 0,
      totalRows: 0,
    };
  }

  const columnWidths: number[] = [];
  for (let c = 1; c <= totalColumns; c++) {
    const col = worksheet.getColumn(c);
    columnWidths.push(col.width ?? DEFAULT_COLUMN_WIDTH);
  }

  const rows: CellData[][] = [];
  const rowHeights: number[] = [];

  for (let r = 1; r <= totalRows; r++) {
    const row = worksheet.getRow(r);
    rowHeights.push(row.height ?? DEFAULT_ROW_HEIGHT);

    const rowData: CellData[] = [];
    for (let c = 1; c <= totalColumns; c++) {
      const mergeInfo = mergeMap.get(`${r},${c}`);
      const isMaster =
        mergeInfo !== undefined &&
        mergeInfo.masterRow === r &&
        mergeInfo.masterCol === c;
      const isSlave = mergeInfo !== undefined && !isMaster;

      const cell = row.getCell(c);
      const value = isSlave ? "" : extractCellValue(cell);
      const style = isSlave ? {} : mapExcelStyle(cell);

      rowData.push({
        value,
        style,
        colSpan: isMaster ? mergeInfo.colSpan : 1,
        rowSpan: isMaster ? mergeInfo.rowSpan : 1,
        isMergedSlave: isSlave,
      });
    }

    rows.push(rowData);
  }

  return {
    name: worksheet.name,
    rows,
    columnWidths,
    rowHeights,
    totalColumns,
    totalRows,
  };
}

export async function readExcel(filePath: string): Promise<SheetData[]> {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const sheets: SheetData[] = [];

  for (const worksheet of workbook.worksheets) {
    const mergeStrings = (worksheet.model?.merges as string[]) ?? [];
    const mergeMap = buildMergeMap(mergeStrings);
    const sheetData = readWorksheet(worksheet, mergeMap);

    if (sheetData.rows.length > 0) {
      sheets.push(sheetData);
    }
  }

  return sheets;
}
