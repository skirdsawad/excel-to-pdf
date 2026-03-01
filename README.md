# excel-to-pdf

CLI tool that converts `.xlsx` files to PDF. Each Excel sheet becomes its own page in the output PDF, preserving cell styles (fonts, colors, borders, fills, alignment, merged cells, column widths).

## Setup

```bash
npm install
npm run build
```

## Usage

```bash
# Build and convert in one step
npm run convert -- input.xlsx
npm run convert -- input.xlsx -o output.pdf

# Or run directly after building
node dist/index.js input.xlsx
node dist/index.js input.xlsx -o output.pdf
```

If no `-o` flag is provided, the output file defaults to the same name with a `.pdf` extension (e.g. `report.xlsx` becomes `report.pdf`).

## What gets preserved

- Font size, bold, italic, underline, strikethrough
- Font color and background fill color
- Cell borders (with color)
- Horizontal alignment (left, center, right, justify)
- Merged cells (colspan and rowspan)
- Column widths (scaled proportionally to fit A4 landscape)
- Multiple sheets (one page per sheet)

## Project structure

```
src/
  index.ts          CLI entry point (commander)
  converter.ts      Orchestrator: read xlsx -> generate PDF
  excel-reader.ts   Read workbook with exceljs, extract data + styles
  pdf-writer.ts     Build pdfmake document definition, write PDF
  style-mapper.ts   Map exceljs styles to pdfmake properties
  types.ts          Shared types and enums
```

## Dependencies

- [exceljs](https://github.com/exceljs/exceljs) - Read .xlsx files with full style information
- [pdfmake](https://github.com/bpampuch/pdfmake) - Generate PDF with declarative table API
- [commander](https://github.com/tj/commander.js) - CLI argument parsing
