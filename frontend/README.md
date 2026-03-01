# ScriptSentries — Frontend

React 18 + TypeScript + Vite + Tailwind CSS

## Setup

```bash
npm install
npm run dev
```

Frontend runs at http://localhost:5173 and proxies `/api` → `http://localhost:8080`

## Key Components

| Component | Description |
|-----------|-------------|
| `App.tsx` | Main layout: Home view + Analyst Workbench. Navbar shows Zero-Retention badge. |
| `UploadZone.tsx` | Drag-and-drop PDF uploader with upload + analysis progress states |
| `RiskTable.tsx` | Paginated, filterable risk table. Columns: Page, Severity, Category, Entity, Status, Redact icon |
| `RiskDrawer.tsx` | Slide-in attorney editor. Shows AI findings, editable fields, **Redaction Toggle** |
| `StatusSelect.tsx` | Color-coded clearance status badge + dropdown |

## Redaction Flow

1. Attorney opens a `RiskFlag` in the Drawer
2. Toggles the "Redact from Export (Security)" switch
3. `PATCH /api/risks/{id}` is called immediately with `{ isRedacted: true }`
4. When "Export Report" is clicked, the backend generates an Excel file with `[REDACTED]` in sensitive columns for those rows
