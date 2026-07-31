import { createWriteStream } from "fs";
import path from "path";
import { Workbook } from "../Workbook";

async function runStreamingExport() {
  const fileTarget = path.join(__dirname, "production_stream_output.xlsx");
  const writeStream = createWriteStream(fileTarget);

  console.log("Orchestrating workbook runtime environment...");
  const workbook = new Workbook(writeStream);

  const sheet = workbook.addWorksheet("Core Ledgers", {
    columns: [{ width: 16 }, { width: 32 }, { width: 18 }, { width: 22 }],
    freezePanes: { rows: 2 },
    autoFilterRange: "A2:D2",
  });
  const bannerStyleId = workbook.styles.registerStyle({
    font: { name: "Segoe UI", size: 14, bold: true, color: "FFFFFFFF" },
    fill: { type: "pattern", patternType: "solid", foregroundColor: "FF1F4E78" },
    alignment: { horizontal: "center", vertical: "center" },
  });

  const headerStyleId = workbook.styles.registerStyle({
    font: { name: "Segoe UI", size: 11, bold: true, color: "FF000000" },
    fill: { type: "pattern", patternType: "solid", foregroundColor: "D9E1F2FF" },
    border: { bottom: { style: "medium", color: "FF000000" } },
  });

  const currencyStyleId = workbook.styles.registerStyle({
    numberFormat: "$#,##0.00",
    font: { name: "Segoe UI", size: 10 },
  });

  sheet.mergeCells("A1:D1");
  sheet.addRow([
    { value: "Enterprise Financial Analytics System Data Pipeline", style: bannerStyleId as any },
  ]);

  sheet.addRow([
    { value: "System Hash ID", style: headerStyleId as any },
    { value: "Corporate Entity Name", style: headerStyleId as any },
    { value: "Revenue Valuation", style: headerStyleId as any },
    { value: "Record Verified Date", style: headerStyleId as any },
  ]);

  console.log("Writing transactional lines...");
  const dataRowCount = 50000;
  for (let i = 1; i <= dataRowCount; i++) {
    const isBackpressured = sheet.addRow([
      `SYS-HEX-${20000 + i}`,
      `Subsidiary Enterprise Division Operations #${i % 5}`,
      { value: 125000.75 + i * 1.5, style: currencyStyleId as any },
      { value: new Date(2026, 6, 31) },
    ]);

    if (isBackpressured) {
      await new Promise<void>((resolve) => {
        sheet["entryStream"].once("drain", () => resolve());
      });
    }
  }

  console.log("Flushing system and committing file payload structure...");
  await workbook.commit();
  console.log(`Pipeline finished successfully. Output written to ${fileTarget}`);
}

process.on("unhandledRejection", (reason) =>
  console.error("💥 CRITICAL UNHANDLED REJECTION:", reason)
);
process.on("uncaughtException", (error) => console.error("💥 CRITICAL UNCAUGHT EXCEPTION:", error));

runStreamingExport().catch((err) => console.error("💥 RUN EXPORT FAILED:", err));
