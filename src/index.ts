import fs from "node:fs";
import path from "node:path";
import { Command } from "commander";
import { convert } from "./converter.js";

const XLSX_EXTENSION = ".xlsx";

const program = new Command();

program
  .name("excel-to-pdf")
  .description("Convert .xlsx files to PDF")
  .version("1.0.0")
  .argument("<input>", "Path to the .xlsx file")
  .option("-o, --output <path>", "Output PDF file path")
  .action(handleConvert);

function handleConvert(input: string, options: { output?: string }): void {
  const inputPath = path.resolve(input);

  if (!fs.existsSync(inputPath)) {
    console.error(`Error: File not found: ${inputPath}`);
    process.exit(1);
  }

  if (!inputPath.toLowerCase().endsWith(XLSX_EXTENSION)) {
    console.error("Error: Input file must be a .xlsx file.");
    process.exit(1);
  }

  const outputPath = options.output
    ? path.resolve(options.output)
    : inputPath.replace(/\.xlsx$/i, ".pdf");

  convert(inputPath, outputPath).catch((err: unknown) => {
    console.error("Conversion failed:", err);
    process.exit(1);
  });
}

program.parse();
