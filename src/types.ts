export enum HorizontalAlignment {
  Left = "left",
  Center = "center",
  Right = "right",
  Justify = "justify",
}

export enum VerticalAlignment {
  Top = "top",
  Middle = "middle",
  Bottom = "bottom",
}

export enum BorderWeight {
  None = 0,
  Thin = 0.5,
  Medium = 1,
  Thick = 2,
}

export interface CellBorder {
  weight: BorderWeight;
  color: string;
}

export interface CellBorders {
  top?: CellBorder;
  bottom?: CellBorder;
  left?: CellBorder;
  right?: CellBorder;
}

export interface CellStyle {
  fontSize?: number;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  fontColor?: string;
  backgroundColor?: string;
  horizontalAlignment?: HorizontalAlignment;
  verticalAlignment?: VerticalAlignment;
  borders?: CellBorders;
  wrapText?: boolean;
}

export interface CellData {
  value: string;
  style: CellStyle;
  colSpan: number;
  rowSpan: number;
  isMergedSlave: boolean;
}

export interface SheetData {
  name: string;
  rows: CellData[][];
  columnWidths: number[];
  rowHeights: number[];
  totalColumns: number;
  totalRows: number;
}

export interface MergeInfo {
  masterRow: number;
  masterCol: number;
  rowSpan: number;
  colSpan: number;
}
