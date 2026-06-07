Warehouse Register Automation
AI-powered web app that converts handwritten warehouse register images into structured Excel reports.

Run & Operate
pnpm --filter @workspace/api-server run dev — run the API server (port 8080)
pnpm --filter @workspace/warehouse-app run dev — run the frontend (port 24630)
pnpm run typecheck — full typecheck across all packages
pnpm run build — typecheck + build all packages
pnpm --filter @workspace/api-spec run codegen — regenerate API hooks and Zod schemas from the OpenAPI spec
Required Environment Variables / Secrets
Key	Where to set	Purpose
GEMINI_API_KEY	Replit Secrets tab (🔒 padlock icon in left sidebar)	Google Gemini AI for image extraction
DATABASE_URL	Replit Secrets tab	Postgres connection string (if DB is added)
To view or change your Gemini API key:

Click the 🔒 Secrets icon in the left sidebar of Replit
Find GEMINI_API_KEY
Click the eye icon to reveal or the edit icon to change it
Get a new key at: https://aistudio.google.com/apikey (free)
Stack
pnpm workspaces, Node.js 24, TypeScript 5.9
Frontend: React + Vite + TailwindCSS
API: Express 5 (Node.js)
AI: Google Gemini 2.5 Flash (via REST API)
Excel: ExcelJS
Validation: Zod, drizzle-zod
API codegen: Orval (from OpenAPI spec)
Build: esbuild (CJS bundle)
Where things live
artifacts/
  api-server/          ← Express backend
    src/routes/
      warehouse.ts     ← Main logic: AI extraction + Excel generation
      health.ts        ← Health check
  warehouse-app/       ← React frontend
    src/App.tsx        ← Single-page UI
lib/
  api-spec/
    openapi.yaml       ← API contract (source of truth)
  api-zod/             ← Generated Zod schemas
  api-client-react/    ← Generated React Query hooks
  db/                  ← Drizzle ORM schema

Core Workflow
User uploads images + selects Inward/Outward + sets opening capacity
  ↓
Backend sends each image to Gemini 2.5 Flash (REST API)
  ↓
AI extracts: date, bond_bags, quantities, location
  ↓
Entries merged by (date + type) as composite key
  → Same date + same type from multiple images → merged into one row
  → Same date + different type → separate rows
  → Location preserved from whichever image has it
  ↓
Grouped by Financial Month (11th → 10th cycle)
  ↓
Excel workbook generated (one sheet per F-Month)
  ↓
User downloads .xlsx

Excel Output Format
Columns: Date | Bags | MT | Capacity | Type | Bond/Bags

Bags — formula: =300+301+270 (sum of numbers after "/" in bond/bags)
MT — formula: =(18225+7895)/1000
Capacity — formula: =D_prev + IF(E_row="Inward", C_row, -C_row)
Bond/Bags — each entry on its own line within the cell
One worksheet per Financial Month
All formulas preserved (never hard-coded values)
Financial Month Rules
11 Feb → 10 Mar  = "Feb-Mar-YYYY"
11 Mar → 10 Apr  = "Mar-Apr-YYYY"
11 Apr → 10 May  = "Apr-May-YYYY"
...

Architecture decisions
Direct Gemini REST API via fetch (not SDK) — avoids esbuild externalization issues with @google/genai
Composite merge key date::type — allows same date to have both inward and outward rows
responseMimeType: "application/json" on Gemini call — forces clean JSON, no markdown fences
ExcelJS formula objects { formula: "...", result: 0 } — Excel recalculates on open, values never stored
Financial month logic: day >= 11 → current month is start; day < 11 → previous month is start
User preferences
Response style: short, direct
No authentication required
Mobile-first simple UI
Gotchas
@google/genai and @google/* packages are externalized in build.mjs — always use direct fetch to Gemini REST API instead of the SDK
ExcelJS cross-sheet formula references use 'SheetName'!CellRef syntax
Pino logger: always use req.log in route handlers, never console.log
