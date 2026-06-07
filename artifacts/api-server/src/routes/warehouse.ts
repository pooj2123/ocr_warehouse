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

// Financial month: 11th of month N → 10th of month N+1
// Returns e.g. "Feb-Mar-2026"
function getFMonth(dateStr: string): string {
  const parts = dateStr.split("-");
  if (parts.length !== 3) return "Unknown";
  const day = parseInt(parts[0]!);
  const month = parseInt(parts[1]!) - 1; // 0-indexed
  const year = parseInt(parts[2]!);
  const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

  let startMonth: number, startYear: number, endMonth: number, endYear: number;
  if (day >= 11) {
    startMonth = month;
    startYear = year;
    endMonth = (month + 1) % 12;
    endYear = month === 11 ? year + 1 : year;
  } else {
    startMonth = (month - 1 + 12) % 12;
    startYear = month === 0 ? year - 1 : year;
    endMonth = month;
    endYear = year;
  }
  return `${monthNames[startMonth]}-${monthNames[endMonth]}-${endYear}`;
}

function parseDate(dateStr: string): Date {
  const parts = dateStr.split("-");
  return new Date(parseInt(parts[2]!), parseInt(parts[1]!) - 1, parseInt(parts[0]!));
}

// Style helpers
function applyBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top: { style: "thin", color: { argb: "FFBDBDBD" } },
    left: { style: "thin", color: { argb: "FFBDBDBD" } },
    bottom: { style: "thin", color: { argb: "FFBDBDBD" } },
    right: { style: "thin", color: { argb: "FFBDBDBD" } },
  };
}

router.post("/process", upload.array("images", 50), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      res.status(400).json({ error: "No images uploaded" });
      return;
    }

    const entryType = (req.body.entryType as string) || "inward";
    const openingCapacityMode = (req.body.openingCapacityMode as string) || "manual";
    const openingCapacityValue = req.body.openingCapacityValue ? parseFloat(req.body.openingCapacityValue) : 0;
    const typeLabel = entryType === "inward" ? "Inward" : "Outward";

    // ── 1. Extract data from each image ──────────────────────────────────
    const allEntries: ExtractedEntry[] = [];

    for (const file of files) {
      const base64 = file.buffer.toString("base64");
      const mimeType = file.mimetype || "image/jpeg";

      const prompt = `Extract all handwritten warehouse register entries from this image.
Return a JSON array. Each element represents one date group:
[
  {
    "date": "DD-MM-YYYY",
    "bond_bags": ["775/300", "759/301"],
    "quantities": [18225, 7895]
  }
]
Rules:
- date must be DD-MM-YYYY format
- bond_bags: array of strings, each in "bondNumber/bagCount" format
- quantities: array of numeric values (weights in KG)
- Group all entries for the same date into one object
- Return [] if nothing found
- Output only JSON, no explanation`;

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
        config: {
          maxOutputTokens: 8192,
          responseMimeType: "application/json",
        },
      });

      const text = response.text ?? "";
      try {
        let parsed: unknown;
        try {
          parsed = JSON.parse(text);
        } catch {
          const match = text.match(/\[[\s\S]*\]/);
          if (match) parsed = JSON.parse(match[0]);
          else throw new Error("No JSON array in response");
        }
        if (Array.isArray(parsed)) {
          allEntries.push(...(parsed as ExtractedEntry[]));
        }
      } catch (parseErr) {
        req.log?.warn({ snippet: text.slice(0, 300), parseErr }, "Failed to parse AI response");
      }
    }

    if (allEntries.length === 0) {
      res.status(400).json({ error: "No data could be extracted from the uploaded images. Please ensure the images are clear and contain handwritten register entries." });
      return;
    }

    // ── 2. Merge by date ─────────────────────────────────────────────────
    const dateMap = new Map<string, ExtractedEntry>();
    for (const entry of allEntries) {
      if (!entry.date || !Array.isArray(entry.bond_bags) || !Array.isArray(entry.quantities)) continue;
      if (dateMap.has(entry.date)) {
        const ex = dateMap.get(entry.date)!;
        ex.bond_bags.push(...entry.bond_bags);
        ex.quantities.push(...entry.quantities);
      } else {
        dateMap.set(entry.date, { ...entry, bond_bags: [...entry.bond_bags], quantities: [...entry.quantities] });
      }
    }

    // ── 3. Group by F-Month ──────────────────────────────────────────────
    const fMonthMap = new Map<string, ExtractedEntry[]>();
    for (const [, entry] of dateMap) {
      const fm = getFMonth(entry.date);
      if (!fMonthMap.has(fm)) fMonthMap.set(fm, []);
      fMonthMap.get(fm)!.push(entry);
    }
    for (const entries of fMonthMap.values()) {
      entries.sort((a, b) => parseDate(a.date).getTime() - parseDate(b.date).getTime());
    }
    const sortedFMonths = Array.from(fMonthMap.keys()).sort((a, b) => {
      const aD = fMonthMap.get(a)![0]!.date;
      const bD = fMonthMap.get(b)![0]!.date;
      return parseDate(aD).getTime() - parseDate(bD).getTime();
    });

    // ── 4. Build Excel workbook ──────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Warehouse Register System";
    workbook.created = new Date();

    // Track last capacity cell reference for carry-forward between sheets
    const sheetClosingRef = new Map<string, { sheet: string; cell: string }>();

    for (let si = 0; si < sortedFMonths.length; si++) {
      const fMonth = sortedFMonths[si]!;
      const entries = fMonthMap.get(fMonth)!;
      const ws = workbook.addWorksheet(fMonth);

      // Column definitions
      // A=Date  B=Type  C=Bond/Bags  D=Bags  E=MT  F=Capacity
      ws.columns = [
        { key: "date",     width: 14 },
        { key: "type",     width: 10 },
        { key: "bondbags", width: 28 },
        { key: "bags",     width: 10 },
        { key: "mt",       width: 12 },
        { key: "capacity", width: 14 },
      ];

      // ── Header row (row 1) ────────────────────────────────────────────
      const hRow = ws.addRow(["Date", "Type", "Bond/Bags", "Bags", "MT", "Capacity"]);
      hRow.eachCell((cell) => {
        cell.font = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2F5496" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        applyBorder(cell);
      });
      hRow.height = 18;

      // ── Opening row (row 2) ───────────────────────────────────────────
      const openingRow = ws.addRow(["Opening", "", "", "", "", null]);
      const openingCapCell = openingRow.getCell(6); // F2

      if (si === 0) {
        // First sheet — use manual value
        openingCapCell.value = openingCapacityValue;
      } else if (openingCapacityMode === "carryforward") {
        // Carry forward from previous sheet's last capacity cell
        const prevFMonth = sortedFMonths[si - 1]!;
        const ref = sheetClosingRef.get(prevFMonth);
        if (ref) {
          // Cross-sheet reference in Excel
          openingCapCell.value = { formula: `'${ref.sheet}'!${ref.cell}`, result: 0 };
        } else {
          openingCapCell.value = openingCapacityValue;
        }
      } else {
        openingCapCell.value = openingCapacityValue;
      }
      openingCapCell.numFmt = '#,##0.000';

      openingRow.eachCell((cell, colNum) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCE6F1" } };
        if (colNum === 1) cell.font = { italic: true };
        applyBorder(cell);
      });

      // Track the previous capacity cell for running balance formulas
      let prevCapRow = 2; // row number of opening

      // ── Data rows ─────────────────────────────────────────────────────
      for (let ei = 0; ei < entries.length; ei++) {
        const entry = entries[ei]!;
        const rowNum = ws.rowCount + 1;

        // Bond/Bags: concatenate raw strings (e.g. "775/300759/301")
        const bondBagsStr = entry.bond_bags.join("");

        // Bags formula: sum of numbers AFTER "/" in each bond_bag
        const bagNumbers = entry.bond_bags.map((bb) => {
          const slash = bb.lastIndexOf("/");
          return slash >= 0 ? bb.slice(slash + 1) : bb;
        });
        const bagsFormula = bagNumbers.join("+");

        // MT formula: =(q1+q2+...)/1000
        const qtyFormula = entry.quantities.length > 0
          ? `(${entry.quantities.join("+")}) / 1000`
          : "0";

        // Capacity formula:
        // Inward:  prev_capacity + this_MT   → =F{prev}+E{row}
        // Outward: prev_capacity - this_MT   → =F{prev}-E{row}
        const capFormula = entryType === "inward"
          ? `F${prevCapRow}+E${rowNum}`
          : `F${prevCapRow}-E${rowNum}`;

        const dataRow = ws.addRow([entry.date, typeLabel, bondBagsStr, null, null, null]);

        // D = Bags
        const bagsCell = dataRow.getCell(4);
        bagsCell.value = { formula: `=${bagsFormula}`, result: 0 };
        bagsCell.numFmt = "#,##0";

        // E = MT
        const mtCell = dataRow.getCell(5);
        mtCell.value = { formula: `=${qtyFormula}`, result: 0 };
        mtCell.numFmt = "#,##0.000";

        // F = Capacity
        const capCell = dataRow.getCell(6);
        capCell.value = { formula: `=${capFormula}`, result: 0 };
        capCell.numFmt = "#,##0.000";

        dataRow.eachCell((cell, colNum) => {
          cell.alignment = { vertical: "middle", horizontal: colNum >= 4 ? "right" : "left" };
          // Alternating row background
          if ((ei + 1) % 2 === 0) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F8FF" } };
          }
          applyBorder(cell);
        });

        prevCapRow = rowNum;
      }

      // ── Closing row ───────────────────────────────────────────────────
      const lastDataRowNum = ws.rowCount;
      const closingRow = ws.addRow(["Closing", "", "", "", "", null]);
      const closingCapCell = closingRow.getCell(6);
      // Just reference the last data row's capacity
      closingCapCell.value = { formula: `=F${lastDataRowNum}`, result: 0 };
      closingCapCell.numFmt = "#,##0.000";
      closingCapCell.font = { bold: true };

      closingRow.eachCell((cell, colNum) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2EFDA" } };
        if (colNum === 1 || colNum === 6) cell.font = { bold: true };
        applyBorder(cell);
      });

      // Save the closing capacity cell reference for carry-forward
      const closingRowNum = closingRow.number;
      sheetClosingRef.set(fMonth, { sheet: fMonth, cell: `F${closingRowNum}` });

      // Freeze top row
      ws.views = [{ state: "frozen", ySplit: 1 }];
    }

    // ── 5. Send response ─────────────────────────────────────────────────
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
