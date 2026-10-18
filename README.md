# Fieldnotes — Sales Insight Dashboard

A responsive, client-side dashboard for exploring sales data from CSV, XLS, and XLSX files. It starts with the included dataset. All parsing, filtering, metrics, charts, and exports run in your browser; uploaded files are not sent to a server.

## Run locally

Requirements: Node.js 22.12 or newer.

```sh
npm install
npm run dev
```

Open the local URL printed by Vite. To create and preview a production build:

```sh
npm run build
npm run preview
```

## Included data

The starter data is `public/sales-data.csv`. To analyze another file, choose **Replace dataset** in the dashboard. CSV, XLS, and XLSX are supported; spreadsheets use the first worksheet and first row as column headers.

Column roles are suggested from header names and can be changed in Settings. Missing roles remain unavailable instead of being estimated. The default status scope includes every source row; use the status filter to inspect completed, pending, returned, or cancelled orders separately.

## Privacy and calculations

- Data processing stays in the browser. The app has no backend, database, integration, or live business-system connection.
- Revenue and profit are sums of the mapped source columns. Order count uses distinct mapped order IDs when available.
- Data quality checks report missing values, exact duplicate rows, and invalid mapped date or numeric values without silently changing the source.
- CSV exports preserve source columns and values for the current filtered view.

## Stack

React, TypeScript, Vite, Tailwind CSS, Recharts, Lucide React, PapaParse, SheetJS (`xlsx`), and Wouter.
