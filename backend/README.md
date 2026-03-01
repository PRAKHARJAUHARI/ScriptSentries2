# ScriptSentries — Backend

Java 21 + Spring Boot 3.3 + Spring AI + PDFBox + PostgreSQL

## Setup

1. PostgreSQL: `CREATE DATABASE scriptsentries;`
2. Set env vars:
   ```
   OPENAI_API_KEY=sk-...
   DB_HOST=localhost
   DB_PASSWORD=yourpassword
   ```
3. Run: `mvn spring-boot:run`

## Key Endpoints

| Method | Path | Description |
|--------|------|-------------|
| POST | `/api/scripts/scan` | Upload PDF for zero-retention AI analysis |
| GET | `/api/scripts` | List all analyzed scripts |
| GET | `/api/scripts/{id}` | Get script + risk flags |
| PATCH | `/api/risks/{id}` | Update status, comments, restrictions, isRedacted |
| GET | `/api/scripts/{id}/export` | Download redacted Excel report |

## Zero-Retention Policy

The PDF is written to a `java.io.File.createTempFile()` location, processed, and deleted inside a `try-finally` block. The raw PDF content is **never** persisted to the database. Only the filename and metadata are stored.

## Security Architecture

- `isRedacted` flag on `RiskFlag` entity
- When `isRedacted=true`, Excel export replaces `entityName`, `snippet`, `comments`, `restrictions` with `[REDACTED]`
- `reason` and `suggestion` remain visible (non-sensitive)
