# ScriptSentries

**AI-Powered Legal Clearance Platform for Hollywood Studios**

## Project Structure

```
scriptsentries/
├── backend/          ← Java 21 + Spring Boot 3.3 + Spring AI + PDFBox + PostgreSQL
│   ├── pom.xml
│   ├── src/main/java/com/scriptsentries/
│   │   ├── model/enums/   RiskCategory, RiskSubCategory, ClearanceStatus, RiskSeverity
│   │   ├── model/         Script.java, RiskFlag.java (with isRedacted security field)
│   │   ├── service/       PdfExtractionService, ScriptAnalysisService, ExcelExportService
│   │   ├── controller/    ScriptController (zero-retention + redacted export)
│   │   ├── dto/           Request/Response DTOs
│   │   └── config/        AiConfig, SecurityConfig
│   └── src/main/resources/application.yml
│
└── frontend/         ← React 18 + TypeScript + Vite + Tailwind CSS
    ├── src/
    │   ├── api/api.ts         All API calls (upload, update, export)
    │   ├── types/index.ts     TypeScript interfaces
    │   └── components/
    │       ├── App.tsx            Main workbench layout
    │       ├── UploadZone.tsx     PDF drag-and-drop with progress
    │       ├── RiskTable.tsx      Filterable risk table
    │       ├── RiskDrawer.tsx     Attorney editor + Redaction Toggle
    │       └── StatusSelect.tsx   Color-coded status dropdown
    └── package.json

```

## Quick Start

### Backend
```bash
cd backend
# Set environment variables:
export OPENAI_API_KEY=sk-your-key
export DB_PASSWORD=yourpassword

mvn spring-boot:run
```

### Frontend
```bash
cd frontend
npm install
npm run dev
```

Then open: http://localhost:5173

## Security Features

| Feature | Implementation |
|---------|---------------|
| **Zero Retention** | PDF deleted from disk in `try-finally` — never stored in DB |
| **Redaction** | `isRedacted` field on `RiskFlag` — masks 4 columns in Excel export |
| **Secure Export** | `[REDACTED]` replaces entityName, snippet, comments, restrictions |
| **CORS** | Configured for localhost dev + production domain |

## AI Analysis

Uses **GPT-4o** via Spring AI with a detailed Contextual Sentiment Analysis prompt:

- Hero uses brand → `LOW / PERMISSIBLE`
- Villain uses brand as weapon → `HIGH / PRODUCT_MISUSE`
- Real person mocked → `HIGH / LIKENESS`
- Song lyrics quoted → `HIGH / MUSIC_CHOREOGRAPHY`
- Phone numbers detected → `MEDIUM / NAMES_NUMBERS`

14 risk categories × 50+ sub-categories mapped to studio taxonomy.
