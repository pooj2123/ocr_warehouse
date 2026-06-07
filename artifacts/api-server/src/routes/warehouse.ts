import { Router } from "express";
import multer from "multer";
import { GoogleGenAI } from "@google/genai";
import ExcelJS from "exceljs";

const router = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

if (!process.env.GEMINI_API_KEY) {
  throw new Error("GEMINI_API_KEY must be set.");
}

const ai = new GoogleGenAI({ apiKey: process.env.GEMINI_API_KEY });

interface ExtractedEntry {
  date: string;
  bond_bags: string[];
  quantities: number[];
}

function getFMonth(dateStr: string): string {
  // Parse date in format DD-MM-YYYY
  const parts = dateStr.split("-");
  if (parts.length !== 3) return "Unknown";
  const day = parseInt(parts[0]!);
  const month = parseInt(parts[1]!) - 1; // 0-indexed
  const year = parseInt(parts[2]!);

  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  let fStartMonth: number;
  let fStartYear: number;
  let fEndMonth: number;
  let fEndYear: number;

  if (day >= 11) {
    fStartMonth = month;
    fStartYear = year;
    fEndMonth = (month + 1) % 12;
    fEndYear = month === 11 ? year + 1 : year;
  } else {
    fStartMonth = (month - 1 + 12) % 12;
    fStartYear = month === 0 ? year - 1 : year;
    fEndMonth = month;
    fEndYear = year;
  }

  return `${monthNames[fStartMonth]}-${monthNames[fEndMonth]}-${fEndYear}`;
}

function parseDate(dateStr: string): Date {
  const parts = dateStr.split("-");
  return new Date(parseInt(parts[2]!), parseInt(parts[1]!) - 1, parseInt(parts[0]!));
}

router.post("/process", upload.array("images", 50), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      res.status(400).json({ error: "No images uploaded" });
      return;
    }

    const entryType = req.body.entryType as string;
    const openingCapacityMode = req.body.openingCapacityMode as string;
    const openingCapacityValue = req.body.openingCapacityValue ? parseFloat(req.body.openingCapacityValue) : null;

    if (!entryType || !["inward", "outward"].includes(entryType)) {
      res.status(400).json({ error: "Invalid entry type" });
      return;
    }

    // Extract data from all images
    const allEntries: ExtractedEntry[] = [];

    for (const file of files) {
      const base64 = file.buffer.toString("base64");
      const mimeType = file.mimetype || "image/jpeg";

      const prompt = `You are a warehouse data extraction assistant. Extract all handwritten register entries from this image.

Return a JSON array of objects with this exact structure:
[
  {
    "date": "DD-MM-YYYY",
    "bond_bags": ["401/565", "403/89"],
    "quantities": [16095, 2230]
  }
]

Rules:
- Group entries by date
- date format must be DD-MM-YYYY
- bond_bags is array of strings in format "number/number"
- quantities is array of numbers (weights/MT values without units)
- If multiple dates exist, return one object per date
- If no data found, return empty array []
- Return only valid JSON, no markdown, no explanation`;

      const response = await ai.models.generateContent({
        model: "gemini-2.5-flash",
        contents: [
          {
            role: "user",
            parts: [
              { inlineData: { mimeType, data: base64 } },
              { text: prompt },
            ],
          },
        ],
        config: { maxOutputTokens: 8192 },
      });

      const text = response.text ?? "";
      const cleaned = text.replace(/```json\n?/g, "").replace(/```\n?/g, "").trim();

      try {
        const parsed = JSON.parse(cleaned) as ExtractedEntry[];
        if (Array.isArray(parsed)) {
          allEntries.push(...parsed);
        }
      } catch {
        req.log?.warn({ text }, "Failed to parse AI response");
      }
    }

    if (allEntries.length === 0) {
      res.status(400).json({ error: "No data could be extracted from the images" });
      return;
    }

    // Merge entries by date
    const dateMap = new Map<string, ExtractedEntry>();
    for (const entry of allEntries) {
      if (dateMap.has(entry.date)) {
        const existing = dateMap.get(entry.date)!;
        existing.bond_bags.push(...entry.bond_bags);
        existing.quantities.push(...entry.quantities);
      } else {
        dateMap.set(entry.date, { ...entry, bond_bags: [...entry.bond_bags], quantities: [...entry.quantities] });
      }
    }

    // Group by F-Month
    const fMonthMap = new Map<string, ExtractedEntry[]>();
    for (const [, entry] of dateMap) {
      const fMonth = getFMonth(entry.date);
      if (!fMonthMap.has(fMonth)) {
        fMonthMap.set(fMonth, []);
      }
      fMonthMap.get(fMonth)!.push(entry);
    }

    // Sort entries within each F-Month by date
    for (const [, entries] of fMonthMap) {
      entries.sort((a, b) => parseDate(a.date).getTime() - parseDate(b.date).getTime());
    }

    // Sort F-Months chronologically
    const sortedFMonths = Array.from(fMonthMap.keys()).sort((a, b) => {
      const aEntries = fMonthMap.get(a)!;
      const bEntries = fMonthMap.get(b)!;
      return parseDate(aEntries[0]!.date).getTime() - parseDate(bEntries[0]!.date).getTime();
    });

    // Generate Excel workbook
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Warehouse Register System";
    workbook.created = new Date();

    const isInward = entryType === "inward";

    // Track closing capacity per sheet for carry-forward
    const closingCapacities = new Map<string, string>(); // fMonth -> cell reference or value

    for (let sheetIndex = 0; sheetIndex < sortedFMonths.length; sheetIndex++) {
      const fMonth = sortedFMonths[sheetIndex]!;
      const entries = fMonthMap.get(fMonth)!;

      const sheet = workbook.addWorksheet(fMonth);

      // Column widths
      sheet.columns = [
        { key: "date", width: 15 },
        { key: "inward", width: 16 },
        { key: "outward", width: 16 },
        { key: "capacity", width: 16 },
        { key: "bags", width: 20 },
      ];

      // Header row
      const headerRow = sheet.addRow(["Date", "Inward MT", "Outward MT", "Capacity MT", "Bags"]);
      headerRow.font = { bold: true };
      headerRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFD9E1F2" } };
      headerRow.border = {
        bottom: { style: "medium" },
      };

      // Opening capacity row (row 2)
      let openingCapacityFormula: string;
      let openingCapacityValue_actual: number | string;

      if (openingCapacityMode === "manual" || sheetIndex === 0) {
        if (openingCapacityMode === "carryforward" && sheetIndex > 0) {
          const prevFMonth = sortedFMonths[sheetIndex - 1]!;
          const prevClosingRef = closingCapacities.get(prevFMonth);
          openingCapacityFormula = prevClosingRef ? `=${prevClosingRef}` : "=0";
          openingCapacityValue_actual = openingCapacityFormula;
        } else {
          openingCapacityValue_actual = openingCapacityValue ?? 0;
          openingCapacityFormula = String(openingCapacityValue_actual);
        }
      } else {
        const prevFMonth = sortedFMonths[sheetIndex - 1]!;
        const prevClosingRef = closingCapacities.get(prevFMonth);
        openingCapacityFormula = prevClosingRef ? `=${prevClosingRef}` : "=0";
        openingCapacityValue_actual = openingCapacityFormula;
      }

      const openingRow = sheet.addRow(["Opening", null, null, null, null]);
      openingRow.font = { italic: true, color: { argb: "FF595959" } };

      // Set opening capacity in column D (capacity)
      const openingCapacityCell = openingRow.getCell(4);
      if (typeof openingCapacityValue_actual === "number") {
        openingCapacityCell.value = openingCapacityValue_actual;
      } else if (openingCapacityFormula.startsWith("=")) {
        openingCapacityCell.value = { formula: openingCapacityFormula.slice(1), result: 0 };
      } else {
        openingCapacityCell.value = openingCapacityValue ?? 0;
      }

      const openingRowNumber = openingRow.number;

      // Data rows
      let lastCapacityRef = `D${openingRowNumber}`;

      for (const entry of entries) {
        const rowNum = sheet.rowCount + 1;

        const row = sheet.addRow([entry.date, null, null, null, null]);

        // Bags formula - only use the number after "/"
        const bagParts = entry.bond_bags.map((bb) => {
          const parts = bb.split("/");
          return parts.length >= 2 ? parts[parts.length - 1]! : parts[0]!;
        });
        const bagsFormula = bagParts.join("+");
        const bagsCell = row.getCell(5);
        bagsCell.value = { formula: `=${bagsFormula}`, result: 0 };

        // MT formula = (sum of quantities) / 1000
        const qtySumStr = entry.quantities.join("+");
        const mtFormula = `=(${qtySumStr})/1000`;

        if (isInward) {
          // Inward MT in column B
          const inwardCell = row.getCell(2);
          inwardCell.value = { formula: mtFormula.slice(1), result: 0 };
          // Outward blank
          row.getCell(3).value = null;
        } else {
          // Outward MT in column C
          const outwardCell = row.getCell(3);
          outwardCell.value = { formula: mtFormula.slice(1), result: 0 };
          // Inward blank
          row.getCell(2).value = null;
        }

        // Running capacity = previous capacity + inward - outward
        const capacityCell = row.getCell(4);
        const capacityFormula = `${lastCapacityRef}+B${rowNum}-C${rowNum}`;
        capacityCell.value = { formula: capacityFormula, result: 0 };
        lastCapacityRef = `D${rowNum}`;

        // Number format
        row.getCell(2).numFmt = "#,##0.000";
        row.getCell(3).numFmt = "#,##0.000";
        row.getCell(4).numFmt = "#,##0.000";
        row.getCell(5).numFmt = "#,##0";

        // Alternating row color
        if (rowNum % 2 === 0) {
          row.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F5F5" } };
        }
      }

      // Store closing capacity reference for carry-forward
      closingCapacities.set(fMonth, lastCapacityRef);

      // Closing row
      const closingRow = sheet.addRow(["Closing", null, null, null, null]);
      closingRow.font = { bold: true };
      closingRow.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2EFDA" } };
      const closingCapacityCell = closingRow.getCell(4);
      closingCapacityCell.value = { formula: lastCapacityRef.replace("D", "D"), result: 0 };
      closingCapacityCell.numFmt = "#,##0.000";

      // Add borders to all data cells
      sheet.eachRow((row, rowNumber) => {
        if (rowNumber >= 1) {
          row.eachCell((cell) => {
            cell.border = {
              top: { style: "thin" },
              left: { style: "thin" },
              bottom: { style: "thin" },
              right: { style: "thin" },
            };
          });
        }
      });
    }

    // Write to buffer
    const buffer = await workbook.xlsx.writeBuffer();

    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="warehouse-register.xlsx"`);
    res.send(Buffer.from(buffer));
  } catch (err) {
    req.log?.error({ err }, "Failed to process register");
    res.status(500).json({ error: "Failed to process register images" });
  }
});

export default router;
