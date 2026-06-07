import { Router } from "express";
import multer from "multer";
import ExcelJS from "exceljs";

const router = Router();
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 20 * 1024 * 1024 } });

if (!process.env.GEMINI_API_KEY) {
  throw new Error("GEMINI_API_KEY must be set.");
}

const GEMINI_API_KEY = process.env.GEMINI_API_KEY;
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=${GEMINI_API_KEY}`;

async function extractFromImage(base64: string, mimeType: string, prompt: string): Promise<string> {
  const body = {
    contents: [
      {
        parts: [
          { inline_data: { mime_type: mimeType, data: base64 } },
          { text: prompt },
        ],
      },
    ],
    generationConfig: {
      response_mime_type: "application/json",
      max_output_tokens: 8192,
    },
  };

  const res = await fetch(GEMINI_URL, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body),
  });

  if (!res.ok) {
    const errText = await res.text();
    throw new Error(`Gemini API error ${res.status}: ${errText.slice(0, 200)}`);
  }

  const data = await res.json() as {
    candidates?: Array<{ content?: { parts?: Array<{ text?: string }> } }>;
  };

  return data?.candidates?.[0]?.content?.parts?.[0]?.text ?? "";
}

interface ExtractedEntry {
  date: string;
  bond_bags: string[];
  quantities: number[];
  location?: string;
}

// Composite key: date + type so same date can have separate inward & outward rows
function entryKey(date: string, type: string) {
  return `${date}::${type}`;
}

// Financial month: 11th of month N → 10th of month N+1, e.g. "Feb-Mar-2026"
function getFMonth(dateStr: string): string {
  const parts = dateStr.split("-");
  if (parts.length !== 3) return "Unknown";
  const day = parseInt(parts[0]!);
  const month = parseInt(parts[1]!) - 1;
  const year = parseInt(parts[2]!);
  const names = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];

  let sm: number, sy: number, em: number, ey: number;
  if (day >= 11) {
    sm = month; sy = year;
    em = (month + 1) % 12; ey = month === 11 ? year + 1 : year;
  } else {
    sm = (month - 1 + 12) % 12; sy = month === 0 ? year - 1 : year;
    em = month; ey = year;
  }
  return `${names[sm]}-${names[em]}-${ey}`;
}

function parseDate(dateStr: string): Date {
  const [d, m, y] = dateStr.split("-");
  return new Date(parseInt(y!), parseInt(m!) - 1, parseInt(d!));
}

function applyBorder(cell: ExcelJS.Cell) {
  cell.border = {
    top:    { style: "thin", color: { argb: "FFBDBDBD" } },
    left:   { style: "thin", color: { argb: "FFBDBDBD" } },
    bottom: { style: "thin", color: { argb: "FFBDBDBD" } },
    right:  { style: "thin", color: { argb: "FFBDBDBD" } },
  };
}

router.post("/process", upload.array("images", 50), async (req, res) => {
  try {
    const files = req.files as Express.Multer.File[];
    if (!files || files.length === 0) {
      res.status(400).json({ error: "No images uploaded" });
      return;
    }

    const entryType       = (req.body.entryType as string) || "inward";
    const capMode         = (req.body.openingCapacityMode as string) || "manual";
    const openingCapValue = req.body.openingCapacityValue ? parseFloat(req.body.openingCapacityValue) : 0;
    const typeLabel       = entryType === "inward" ? "Inward" : "Outward";

    // ── 1. Extract from every image ──────────────────────────────────────
    const allEntries: ExtractedEntry[] = [];

    const prompt = `Extract all handwritten warehouse register entries from this image.
Return a JSON array. Each element is one date group:
[
  {
    "date": "DD-MM-YYYY",
    "bond_bags": ["775/300", "759/301"],
    "quantities": [18225, 7895],
    "location": "Shed A"
  }
]
Rules:
- date: DD-MM-YYYY format only
- bond_bags: array of strings in "bondNumber/bagCount" format
- quantities: array of numeric weights (KG)
- location: string if a location/shed is written, otherwise omit the field
- Group all entries for the SAME date into ONE object
- Return [] if nothing found
- Output only the JSON array, no markdown, no explanation`;

    for (const file of files) {
      const base64  = file.buffer.toString("base64");
      const mimeType = file.mimetype || "image/jpeg";
      const text    = await extractFromImage(base64, mimeType, prompt);

      try {
        let parsed: unknown;
        try {
          parsed = JSON.parse(text);
        } catch {
          const match = text.match(/\[[\s\S]*\]/);
          parsed = match ? JSON.parse(match[0]) : null;
        }
        if (Array.isArray(parsed)) {
          allEntries.push(...(parsed as ExtractedEntry[]));
        }
      } catch (e) {
        req.log?.warn({ snippet: text.slice(0, 300), e }, "AI parse failed");
      }
    }

    if (allEntries.length === 0) {
      res.status(400).json({ error: "No data could be extracted. Please ensure the images are clear and contain handwritten register entries." });
      return;
    }

    // ── 2. Merge: composite key = date + entryType ────────────────────────
    // Same date, same type → merge (handles multiple images covering same date)
    // Same date, different type → separate row (inward vs outward same day)
    const mergeMap = new Map<string, ExtractedEntry>();

    for (const entry of allEntries) {
      if (!entry.date || !Array.isArray(entry.bond_bags) || !Array.isArray(entry.quantities)) continue;
      const key = entryKey(entry.date, typeLabel);
      if (mergeMap.has(key)) {
        const ex = mergeMap.get(key)!;
        // Merge bond_bags & quantities (avoid exact duplicates in bond_bags)
        for (const bb of entry.bond_bags) {
          if (!ex.bond_bags.includes(bb)) ex.bond_bags.push(bb);
        }
        ex.quantities.push(...entry.quantities);
        // Preserve location from whichever image has it (date = primary key)
        if (!ex.location && entry.location) ex.location = entry.location;
      } else {
        mergeMap.set(key, {
          date:      entry.date,
          bond_bags: [...entry.bond_bags],
          quantities: [...entry.quantities],
          location:  entry.location,
        });
      }
    }

    // ── 3. Group by F-Month ──────────────────────────────────────────────
    const fMonthMap = new Map<string, ExtractedEntry[]>();
    for (const entry of mergeMap.values()) {
      const fm = getFMonth(entry.date);
      if (!fMonthMap.has(fm)) fMonthMap.set(fm, []);
      fMonthMap.get(fm)!.push(entry);
    }
    for (const arr of fMonthMap.values()) {
      arr.sort((a, b) => parseDate(a.date).getTime() - parseDate(b.date).getTime());
    }
    const sortedFMonths = Array.from(fMonthMap.keys()).sort((a, b) => {
      const aD = fMonthMap.get(a)![0]!.date;
      const bD = fMonthMap.get(b)![0]!.date;
      return parseDate(aD).getTime() - parseDate(bD).getTime();
    });

    // ── 4. Build Excel ────────────────────────────────────────────────────
    const workbook = new ExcelJS.Workbook();
    workbook.creator = "Warehouse Register";
    workbook.created = new Date();

    // Column layout (A–F):
    //  A = Date
    //  B = Bags       (formula: =bag1+bag2+...)
    //  C = MT         (formula: =(q1+q2+...)/1000)
    //  D = Capacity   (formula: =D_prev + IF(E_row="Inward", C_row, -C_row))
    //  E = Type       (Inward / Outward)   ← second to last
    //  F = Bond/Bags  (raw strings, one per line)  ← last

    const sheetClosingRef = new Map<string, { sheet: string; cell: string }>();

    for (let si = 0; si < sortedFMonths.length; si++) {
      const fMonth  = sortedFMonths[si]!;
      const entries = fMonthMap.get(fMonth)!;
      const ws      = workbook.addWorksheet(fMonth);

      ws.columns = [
        { key: "date",     width: 14 },
        { key: "bags",     width: 10 },
        { key: "mt",       width: 12 },
        { key: "capacity", width: 14 },
        { key: "type",     width: 10 },
        { key: "bondbags", width: 30 },
      ];

      // ── Header (row 1) ────────────────────────────────────────────────
      const hRow = ws.addRow(["Date", "Bags", "MT", "Capacity", "Type", "Bond/Bags"]);
      hRow.eachCell((cell) => {
        cell.font      = { bold: true, color: { argb: "FFFFFFFF" } };
        cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF2F5496" } };
        cell.alignment = { horizontal: "center", vertical: "middle" };
        applyBorder(cell);
      });
      hRow.height = 18;

      // ── Opening row (row 2) ───────────────────────────────────────────
      const openRow = ws.addRow(["Opening", null, null, null, "", ""]);
      const openCapCell = openRow.getCell(4); // D2

      if (si === 0) {
        openCapCell.value = openingCapValue;
      } else if (capMode === "carryforward") {
        const prev = sheetClosingRef.get(sortedFMonths[si - 1]!);
        openCapCell.value = prev
          ? { formula: `'${prev.sheet}'!${prev.cell}`, result: 0 }
          : openingCapValue;
      } else {
        openCapCell.value = openingCapValue;
      }
      openCapCell.numFmt = "#,##0.000";

      openRow.eachCell((cell, col) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFDCE6F1" } };
        if (col === 1) cell.font = { italic: true };
        applyBorder(cell);
      });

      let prevCapRow = 2; // row number of the opening row

      // ── Data rows ─────────────────────────────────────────────────────
      for (let ei = 0; ei < entries.length; ei++) {
        const entry  = entries[ei]!;
        const rowNum = ws.rowCount + 1;

        // B: Bags formula — sum of numbers after "/" in each bond_bag
        const bagNums    = entry.bond_bags.map((bb) => bb.slice(bb.lastIndexOf("/") + 1));
        const bagsFormula = bagNums.length ? bagNums.join("+") : "0";

        // C: MT formula — =(q1+q2+...)/1000
        const qtySumStr = entry.quantities.length ? entry.quantities.join("+") : "0";
        const mtFormula = `(${qtySumStr})/1000`;

        // D: Capacity — =D_prev + IF(E_row="Inward", C_row, -C_row)
        const capFormula = `D${prevCapRow}+IF(E${rowNum}="Inward",C${rowNum},-C${rowNum})`;

        // F: Bond/Bags — each entry on its own line
        const bondBagsText = entry.bond_bags.join("\n");

        const dataRow = ws.addRow([entry.date, null, null, null, typeLabel, bondBagsText]);

        // B = Bags
        const bagsCell = dataRow.getCell(2);
        bagsCell.value  = { formula: `=${bagsFormula}`, result: 0 };
        bagsCell.numFmt = "#,##0";

        // C = MT
        const mtCell  = dataRow.getCell(3);
        mtCell.value  = { formula: `=${mtFormula}`, result: 0 };
        mtCell.numFmt = "#,##0.000";

        // D = Capacity
        const capCell  = dataRow.getCell(4);
        capCell.value  = { formula: `=${capFormula}`, result: 0 };
        capCell.numFmt = "#,##0.000";

        // F = Bond/Bags with line-wrap
        const bbCell        = dataRow.getCell(6);
        bbCell.alignment    = { wrapText: true, vertical: "top" };

        // Row height: roughly 15pt per line
        const lineCount = entry.bond_bags.length || 1;
        dataRow.height  = Math.max(18, lineCount * 16);

        dataRow.eachCell((cell, col) => {
          if (col !== 6) {
            cell.alignment = { vertical: "middle", horizontal: col >= 2 && col <= 4 ? "right" : "left" };
          }
          if ((ei + 1) % 2 === 0) {
            cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF5F8FF" } };
          }
          applyBorder(cell);
        });

        prevCapRow = rowNum;
      }

      // ── Closing row ───────────────────────────────────────────────────
      const lastDataRow = ws.rowCount;
      const closeRow    = ws.addRow(["Closing", null, null, null, "", ""]);
      const closeCapCell = closeRow.getCell(4);
      closeCapCell.value  = { formula: `=D${lastDataRow}`, result: 0 };
      closeCapCell.numFmt = "#,##0.000";

      closeRow.eachCell((cell, col) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFE2EFDA" } };
        if (col === 1 || col === 4) cell.font = { bold: true };
        applyBorder(cell);
      });

      sheetClosingRef.set(fMonth, { sheet: fMonth, cell: `D${closeRow.number}` });
      ws.views = [{ state: "frozen", ySplit: 1 }];
    }

    // ── 5. Send file ──────────────────────────────────────────────────────
    const buffer = await workbook.xlsx.writeBuffer();
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="warehouse-register.xlsx"`);
    res.send(Buffer.from(buffer));

  } catch (err) {
    req.log?.error({ err }, "Register processing failed");
    res.status(500).json({ error: "Failed to process register images" });
  }
});

export default router;
