import { readExcel } from "./excel-reader.js";
import { generatePdf } from "./pdf-writer.js";

export async function convert(
  inputPath: string,
  outputPath: string
): Promise<void> {
  console.log(`Reading: ${inputPath}`);
  const sheets = await readExcel(inputPath);

  if (sheets.length === 0) {
    console.log("No sheets with data found in the workbook.");

    return;
  }

  console.log(`Found ${sheets.length} sheet(s): ${sheets.map((s) => s.name).join(", ")}`);

  console.log(`Writing PDF: ${outputPath}`);
  await generatePdf(sheets, outputPath);

  console.log("Done.");
}
