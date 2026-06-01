# RASA Weekly Business Review Dashboard

Next.js port of the WBR dashboard. Reads three Excel workbooks per week (WBR / Loyalty / Catering) from `data/<week>/` and renders the full nine-tab dashboard.

## Stack

- **Next.js 16** (App Router) · **React 19** · JavaScript
- **xlsx** (SheetJS) for server-side parsing
- **chart.js** + **react-chartjs-2** for charts
- Plain CSS (`app/globals.css`)

## Run locally

```bash
cd D:/Kutlerri/WBR-Dashboard
npm install
npm run dev
```

Open http://localhost:3000

## How data flows

```
data/<week>/three xlsx files
        │
        ▼
GET /api/sheets             → lists every folder under data/
GET /api/data/[week]        → lib/xlsxParser.js → parsed JSON for that week
POST /api/upload            → saves three uploaded files into data/<week>/
        │
        ▼
app/page.jsx (week dropdown + tab nav)
        │
        ▼
components/Snapshot.jsx, Sales.jsx, Costs.jsx, Reviews.jsx,
ThirdParty.jsx, Bikky.jsx, Loyalty.jsx, Marketing.jsx, CateringSales.jsx
```

The parser library `lib/xlsxParser.js` is the heart of the app. It mirrors the `parseWBR`, `parseLoyalty`, `parseCateringWB` functions from the original HTML dashboard, plus the per-location + YTD extraction logic that the HTML had hardcoded.

## Adding a new week

Two options:

1. **Drag & drop**: Click "⬆ Upload Data" in the header, give it a name (e.g. `25-31`), pick the three xlsx files, click Upload. Files land in `data/25-31/` on the server.
2. **Manually**: Create `data/<week-name>/`, drop the three xlsx files in. Refresh the page — the new week shows up in the dropdown.

The parser auto-detects which file is which by filename substring:
- `weekly review` / `wbr` / `powered by kutlerri` → WBR workbook
- `loyalty` → Loyalty workbook
- `catering` / `internal purpose` → Catering workbook

## Deploy to Vercel

```bash
# from project root
vercel --prod
```

**Important:** Vercel's filesystem is read-only at runtime — uploads via the modal won't persist on Vercel. For production, either:
- Commit `data/<week>/` folders to the repo (they ship at build time), or
- Wire `/api/upload` to S3 / R2 / Vercel Blob storage instead of `fs.writeFileSync`.

Local dev (`npm run dev` or `npm start` on your own server) supports uploads natively.

## File layout

```
WBR-Dashboard/
├── app/
│   ├── api/
│   │   ├── sheets/route.js
│   │   ├── data/[week]/route.js
│   │   └── upload/route.js
│   ├── globals.css
│   ├── layout.jsx
│   └── page.jsx
├── components/
│   ├── Table.jsx
│   ├── UploadModal.jsx
│   ├── Snapshot.jsx       Sales.jsx          Costs.jsx
│   ├── Reviews.jsx        ThirdParty.jsx     Bikky.jsx
│   ├── Loyalty.jsx        Marketing.jsx      CateringSales.jsx
├── lib/
│   ├── xlsxParser.js   ← parses WBR + Loyalty + Catering workbooks
│   ├── api.js          ← client-side fetch helpers
│   ├── chartSetup.js   ← chart.js registration
│   └── fmt.js          ← formatters (fmt$, fmtPct, fmtVar)
├── data/
│   └── 11-17/                    (preloaded — 3 xlsx files)
├── next.config.js
├── package.json
└── jsconfig.json
```

## Source-of-truth mapping

| Tab in UI               | Component             | Comes from XLSX sheet(s)                                                                     |
| ----------------------- | --------------------- | -------------------------------------------------------------------------------------------- |
| Overview                | Snapshot.jsx          | `Weekly/Period Flash Results Sales`, `Restaurants Revenue P vs A`                            |
| Sales & Revenue         | Sales.jsx             | `ALL - Weekly/PTD Revenue Center *`, `Total Revenue Center P v A`, sub-cat sheets, per-loc sheets |
| Costs                   | Costs.jsx             | `Weekly/Period Flash Results COSTS`                                                          |
| Reviews & Ratings       | Reviews.jsx           | `Weekly in-store leadership metr` (in-store + 90-day side-by-side)                           |
| 3rd Party               | ThirdParty.jsx        | `3PD Reporting - UE & DD`                                                                    |
| Customer Insights       | Bikky.jsx             | `Customer Insights` (Locations + Acquisition + Onboarding)                                   |
| Loyalty                 | Loyalty.jsx           | `Lifecycle - Table`, `WoW 2026`, `Instore Orders`, `Digital Orders`                          |
| Marketing               | Marketing.jsx         | `InputsOutputs Catering`, `SMS - Table`                                                      |
| Catering Sales          | CateringSales.jsx     | `Week of <Date>`, `Sheet1`, `WoWComparision`                                                 |
