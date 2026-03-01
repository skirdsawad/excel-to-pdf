import type ExcelJS from "exceljs";
import {
  BorderWeight,
  HorizontalAlignment,
  VerticalAlignment,
  type CellBorder,
  type CellBorders,
  type CellStyle,
} from "./types.js";

const EXCEL_BORDER_WEIGHT_MAP: Record<string, BorderWeight> = {
  thin: BorderWeight.Thin,
  dotted: BorderWeight.Thin,
  hair: BorderWeight.Thin,
  medium: BorderWeight.Medium,
  double: BorderWeight.Medium,
  dashed: BorderWeight.Thin,
  dashDot: BorderWeight.Thin,
  dashDotDot: BorderWeight.Thin,
  slantDashDot: BorderWeight.Medium,
  mediumDashed: BorderWeight.Medium,
  mediumDashDot: BorderWeight.Medium,
  mediumDashDotDot: BorderWeight.Medium,
  thick: BorderWeight.Thick,
};

const EXCEL_HORIZONTAL_ALIGNMENT_MAP: Record<string, HorizontalAlignment> = {
  left: HorizontalAlignment.Left,
  center: HorizontalAlignment.Center,
  right: HorizontalAlignment.Right,
  fill: HorizontalAlignment.Left,
  justify: HorizontalAlignment.Justify,
  centerContinuous: HorizontalAlignment.Center,
  distributed: HorizontalAlignment.Justify,
};

const EXCEL_VERTICAL_ALIGNMENT_MAP: Record<string, VerticalAlignment> = {
  top: VerticalAlignment.Top,
  middle: VerticalAlignment.Middle,
  bottom: VerticalAlignment.Bottom,
  distributed: VerticalAlignment.Middle,
  justify: VerticalAlignment.Middle,
};

export function argbToHex(argb: string | undefined): string | undefined {
  if (!argb) {
    return undefined;
  }

  // ARGB format: "FFRRGGBB" (8 chars) or "RRGGBB" (6 chars)
  if (argb.length === 8) {
    return `#${argb.substring(2)}`;
  }

  if (argb.length === 6) {
    return `#${argb}`;
  }

  return undefined;
}

function mapBorderSide(
  border: Partial<ExcelJS.Border> | undefined
): CellBorder | undefined {
  if (!border || !border.style) {
    return undefined;
  }

  const weight = EXCEL_BORDER_WEIGHT_MAP[border.style] ?? BorderWeight.Thin;
  if (weight === BorderWeight.None) {
    return undefined;
  }

  const color = argbToHex(border.color?.argb) ?? "#000000";

  return { weight, color };
}

function mapBorders(
  borders: Partial<ExcelJS.Borders> | undefined
): CellBorders | undefined {
  if (!borders) {
    return undefined;
  }

  const top = mapBorderSide(borders.top);
  const bottom = mapBorderSide(borders.bottom);
  const left = mapBorderSide(borders.left);
  const right = mapBorderSide(borders.right);

  if (!top && !bottom && !left && !right) {
    return undefined;
  }

  return { top, bottom, left, right };
}

function mapHorizontalAlignment(
  alignment: string | undefined
): HorizontalAlignment | undefined {
  if (!alignment) {
    return undefined;
  }

  return EXCEL_HORIZONTAL_ALIGNMENT_MAP[alignment];
}

function mapVerticalAlignment(
  alignment: string | undefined
): VerticalAlignment | undefined {
  if (!alignment) {
    return undefined;
  }

  return EXCEL_VERTICAL_ALIGNMENT_MAP[alignment];
}

export function mapExcelStyle(cell: ExcelJS.Cell): CellStyle {
  const style: CellStyle = {};
  const font = cell.font;
  const fill = cell.fill;
  const alignment = cell.alignment;
  const border = cell.border;

  if (font) {
    if (font.size) {
      style.fontSize = font.size;
    }
    if (font.bold) {
      style.bold = true;
    }
    if (font.italic) {
      style.italic = true;
    }
    if (font.underline && font.underline !== "none") {
      style.underline = true;
    }
    if (font.strike) {
      style.strikethrough = true;
    }
    if (font.color?.argb) {
      const hex = argbToHex(font.color.argb);
      if (hex) {
        style.fontColor = hex;
      }
    }
  }

  if (fill && fill.type === "pattern" && fill.pattern !== "none") {
    const patternFill = fill as ExcelJS.FillPattern;
    const bgColor = patternFill.fgColor?.argb;
    if (bgColor) {
      const hex = argbToHex(bgColor);
      if (hex) {
        style.backgroundColor = hex;
      }
    }
  }

  if (alignment) {
    const hAlign = mapHorizontalAlignment(alignment.horizontal);
    if (hAlign) {
      style.horizontalAlignment = hAlign;
    }

    const vAlign = mapVerticalAlignment(alignment.vertical);
    if (vAlign) {
      style.verticalAlignment = vAlign;
    }

    if (alignment.wrapText) {
      style.wrapText = true;
    }
  }

  if (border) {
    const mappedBorders = mapBorders(border);
    if (mappedBorders) {
      style.borders = mappedBorders;
    }
  }

  return style;
}
