package com.scriptsentries.service;

import com.scriptsentries.model.*;
import com.scriptsentries.model.enums.ClearanceStatus;
import com.scriptsentries.model.enums.RiskCategory;
import com.scriptsentries.repository.ProjectMemberRepository;
import com.scriptsentries.repository.RiskFlagRepository;
import com.scriptsentries.repository.ScriptRepository;
import com.scriptsentries.repository.UserRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;

/**
 * E&O (Errors & Omissions) Insurance Export Service
 *
 * Produces a 3-tab Excel workbook structured for US E&O insurance submission:
 *
 *   Tab 1 — PRODUCTION COVER SHEET
 *     Production details, attorney certification, risk summary statistics.
 *     Mirrors the standard "Production Information" page required by carriers
 *     such as DeWitt Stern, Front Row Insurance, and Media/Entertainment Group.
 *
 *   Tab 2 — RISK LOG (Full)
 *     Every identified risk with legal exposure category, clearance status,
 *     restrictions, and attorney notes. Redacted rows are masked per attorney
 *     privilege flags. Line numbers provided for insurer cross-reference.
 *
 *   Tab 3 — CLEARED ITEMS SUMMARY
 *     Only cleared/permissible items. The insurer uses this to confirm the
 *     production has addressed all flagged items.
 *
 * US Legal Category Mapping:
 *   The AI uses ScriptSentries internal categories (LIKENESS, MUSIC_CHOREOGRAPHY,
 *   etc.). This service maps those to the standard legal exposure labels that
 *   US E&O underwriters recognise on submission forms.
 */
@Service
@RequiredArgsConstructor
@Slf4j
public class EoExportService {

    private final ScriptRepository        scriptRepo;
    private final RiskFlagRepository      riskFlagRepo;
    private final UserRepository          userRepo;
    private final ProjectMemberRepository memberRepo;

    private static final String REDACTED      = "[REDACTED — ATTORNEY PRIVILEGE]";
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("MM/dd/yyyy");
    private static final DateTimeFormatter DATETIME_FMT = DateTimeFormatter.ofPattern("MM/dd/yyyy HH:mm");

    /**
     * US E&O legal exposure category labels.
     * Maps ScriptSentries internal RiskCategory enum → human-readable US legal label.
     *
     * These labels match the terminology on standard US E&O submission forms.
     */
    private static final Map<String, String> LEGAL_EXPOSURE_LABELS = Map.ofEntries(
            Map.entry("FACT_BASED_ISSUES",    "Defamation / False Light / Invasion of Privacy"),
            Map.entry("GOVERNMENT",           "Government / Political Figure Portrayal"),
            Map.entry("LIKENESS",             "Right of Publicity / Likeness Rights (17 U.S.C. / State Law)"),
            Map.entry("LOCATIONS",            "Location Release / Trademark in Background"),
            Map.entry("MARKETING_ADDED_VALUE","Promotional Tie-in / Product Placement"),
            Map.entry("MUSIC_CHOREOGRAPHY",   "Music Sync Rights / Choreography Copyright (17 U.S.C. § 106)"),
            Map.entry("NAMES_NUMBERS",        "Coincidental Use / Real Person Name / Phone/Address"),
            Map.entry("PLAYBACK",             "On-Screen Playback Copyright (Film/TV/Web Content)"),
            Map.entry("PRODUCT_MISUSE",       "Product Disparagement / Defamation of Goods"),
            Map.entry("PROPS_SET_DRESSING",   "Prop/Set Trademark Clearance / Art Direction IP"),
            Map.entry("REFERENCES",           "Literary/Cultural Reference — Copyright / Trademark"),
            Map.entry("VEHICLES",             "Vehicle Trademark / License Plate / VIN Clearance"),
            Map.entry("WARDROBE",             "Fashion Trademark / Designer Brand Clearance"),
            Map.entry("OTHER",                "Miscellaneous Legal Exposure — See Notes")
    );

    private static final Map<String, String> STATUS_LABELS = Map.of(
            "PENDING",                "⏳ Pending Review",
            "CLEARED",                "✅ Cleared",
            "NOT_CLEAR",              "❌ Not Clear",
            "NEGOTIATED_BY_ATTORNEY", "⚖️ Negotiated by Attorney",
            "BRANDED_INTEGRATION",    "🤝 Branded Integration Agreement",
            "NO_CLEARANCE_NECESSARY", "✅ No Clearance Necessary",
            "PERMISSIBLE",            "✅ Permissible"
    );

    // ─────────────────────────────────────────────────────────────────────────
    // Public API
    // ─────────────────────────────────────────────────────────────────────────

    public byte[] generateEoReport(Long scriptId, Long requestingUserId) throws IOException {
        Script script = scriptRepo.findById(scriptId)
                .orElseThrow(() -> new RuntimeException("Script not found: " + scriptId));

        List<RiskFlag> risks = riskFlagRepo.findByScriptOrderByPageNumberAscSeverityDesc(script);

        // Find certifying attorney (highest-role member of the project)
        String certifyingAttorney = resolveCertifyingAttorney(script, requestingUserId);

        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            Styles s = new Styles(wb);

            buildCoverSheet(wb, s, script, risks, certifyingAttorney);
            buildRiskLog(wb, s, risks);
            buildClearedSummary(wb, s, risks);

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            wb.write(out);

            log.info("E&O report generated for script '{}': {} risks, {} cleared",
                    script.getFilename(), risks.size(),
                    risks.stream().filter(r -> isCleared(r.getStatus())).count());

            return out.toByteArray();
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Tab 1 — Production Cover Sheet
    // ─────────────────────────────────────────────────────────────────────────

    private void buildCoverSheet(XSSFWorkbook wb, Styles s, Script script,
                                 List<RiskFlag> risks, String certifyingAttorney) {
        XSSFSheet sheet = wb.createSheet("E&O Cover Sheet");
        sheet.setColumnWidth(0, 38 * 256);
        sheet.setColumnWidth(1, 55 * 256);

        int row = 0;

        // ── Header banner ──
        row = mergedRow(sheet, row, s.title,
                "ERRORS & OMISSIONS INSURANCE — PRODUCTION CLEARANCE SUBMISSION", 2);
        row = mergedRow(sheet, row, s.subtitle,
                "Prepared by ScriptSentries  |  United States Production", 2);
        row = mergedRow(sheet, row, s.subtitle,
                "Generated: " + LocalDateTime.now().format(DATETIME_FMT), 2);
        row++; // spacer

        // ── Section: Production Information ──
        row = sectionHeader(sheet, row, s.sectionHeader, "PRODUCTION INFORMATION", 2);

        Project p = script.getProject();
        row = labelValue(sheet, row, s.label, s.value, "Production Title",
                p != null ? p.getName() : script.getFilename());
        row = labelValue(sheet, row, s.label, s.value, "Studio / Production Company",
                p != null && p.getStudioName() != null ? p.getStudioName() : "—");
        row = labelValue(sheet, row, s.label, s.value, "Director",
                p != null && p.getDirector() != null ? p.getDirector() : "—");
        row = labelValue(sheet, row, s.label, s.value, "Producer",
                p != null && p.getProducer() != null ? p.getProducer() : "—");
        row = labelValue(sheet, row, s.label, s.value, "Genre",
                p != null && p.getGenre() != null ? p.getGenre() : "—");
        row = labelValue(sheet, row, s.label, s.value, "Expected Release",
                p != null && p.getExpectedRelease() != null ? p.getExpectedRelease() : "—");
        row = labelValue(sheet, row, s.label, s.value, "IMDB / Project Link",
                p != null && p.getImdbLink() != null ? p.getImdbLink() : "—");
        row = labelValue(sheet, row, s.label, s.value, "Production Contact Email",
                p != null && p.getProductionEmail() != null ? p.getProductionEmail() : "—");
        row = labelValue(sheet, row, s.label, s.value, "Production Contact Phone",
                p != null && p.getProductionPhone() != null ? p.getProductionPhone() : "—");
        row++;

        // ── Section: Script Version ──
        row = sectionHeader(sheet, row, s.sectionHeader, "SCRIPT VERSION ANALYZED", 2);
        row = labelValue(sheet, row, s.label, s.value, "Script Filename", script.getFilename());
        row = labelValue(sheet, row, s.label, s.value, "Version Label",
                script.getVersionName() != null ? script.getVersionName() : "—");
        row = labelValue(sheet, row, s.label, s.value, "Total Pages",
                String.valueOf(script.getTotalPages()));
        row = labelValue(sheet, row, s.label, s.value, "Analysis Date",
                script.getUploadedAt() != null ? script.getUploadedAt().format(DATE_FMT) : "—");
        row++;

        // ── Section: Risk Summary ──
        row = sectionHeader(sheet, row, s.sectionHeader, "RISK SUMMARY", 2);

        long high    = risks.stream().filter(r -> "HIGH".equals(r.getSeverity().name())).count();
        long medium  = risks.stream().filter(r -> "MEDIUM".equals(r.getSeverity().name())).count();
        long low     = risks.stream().filter(r -> "LOW".equals(r.getSeverity().name())).count();
        long cleared = risks.stream().filter(r -> isCleared(r.getStatus())).count();
        long pending = risks.stream().filter(r -> r.getStatus() == ClearanceStatus.PENDING).count();
        long notClear = risks.stream().filter(r -> r.getStatus() == ClearanceStatus.NOT_CLEAR).count();
        long redacted = risks.stream().filter(RiskFlag::isRedacted).count();

        row = labelValue(sheet, row, s.label, s.value, "Total Risks Identified", String.valueOf(risks.size()));
        row = labelValue(sheet, row, s.labelHigh, s.valueHigh, "High Severity", String.valueOf(high));
        row = labelValue(sheet, row, s.labelMed, s.valueMed, "Medium Severity", String.valueOf(medium));
        row = labelValue(sheet, row, s.labelLow, s.valueLow, "Low Severity", String.valueOf(low));
        row = labelValue(sheet, row, s.label, s.value, "Cleared / Permissible", String.valueOf(cleared));
        row = labelValue(sheet, row, s.labelPending, s.valuePending, "Pending Review", String.valueOf(pending));
        row = labelValue(sheet, row, s.labelNotClear, s.valueNotClear, "Not Clear", String.valueOf(notClear));
        row = labelValue(sheet, row, s.label, s.value, "Attorney-Privileged (Redacted)", String.valueOf(redacted));
        row++;

        // ── Section: Attorney Certification ──
        row = sectionHeader(sheet, row, s.sectionHeader, "ATTORNEY CERTIFICATION", 2);
        row = labelValue(sheet, row, s.label, s.value, "Certifying Attorney", certifyingAttorney);
        row = labelValue(sheet, row, s.label, s.value, "Certification Date",
                LocalDateTime.now().format(DATE_FMT));
        row = mergedRow(sheet, row, s.certText,
                "I certify that the above-listed risks have been reviewed and addressed " +
                        "in accordance with applicable US copyright, trademark, right of publicity, " +
                        "and defamation law. Cleared items are documented in the attached risk log.", 2);
        row++;

        // ── Section: Legal Framework Note ──
        row = sectionHeader(sheet, row, s.sectionHeader, "APPLICABLE US LEGAL FRAMEWORK", 2);
        row = mergedRow(sheet, row, s.legalNote,
                "Copyright: 17 U.S.C. § 101 et seq. (fair use analysis per 17 U.S.C. § 107)", 2);
        row = mergedRow(sheet, row, s.legalNote,
                "Right of Publicity: State law (California Civil Code § 3344; NY Civil Rights Law §§ 50-51)", 2);
        row = mergedRow(sheet, row, s.legalNote,
                "Trademark: 15 U.S.C. § 1051 et seq. (Lanham Act) — nominative use doctrine applicable", 2);
        row = mergedRow(sheet, row, s.legalNote,
                "Defamation / False Light: New York Times v. Sullivan (1964) actual malice standard", 2);
        row = mergedRow(sheet, row, s.legalNote,
                "Music: 17 U.S.C. §§ 106, 114, 115 — sync license + master use license required", 2);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Tab 2 — Full Risk Log
    // ─────────────────────────────────────────────────────────────────────────

    private void buildRiskLog(XSSFWorkbook wb, Styles s, List<RiskFlag> risks) {
        XSSFSheet sheet = wb.createSheet("Full Risk Log");

        // Column widths
        int[] widths = {8, 10, 14, 28, 40, 20, 40, 45, 28, 40, 40, 12};
        for (int i = 0; i < widths.length; i++) sheet.setColumnWidth(i, widths[i] * 256);

        // Header row
        String[] headers = {
                "Line #", "Page", "Severity", "US Legal Exposure Category",
                "Entity / Item", "Clearance Status", "Script Excerpt",
                "Legal Basis", "Restrictions / Conditions",
                "Attorney Notes", "AI Recommendation", "Privileged"
        };
        int row = 0;
        Row headerRow = sheet.createRow(row++);
        for (int col = 0; col < headers.length; col++) {
            Cell cell = headerRow.createCell(col);
            cell.setCellValue(headers[col]);
            cell.setCellStyle(s.colHeader);
        }
        sheet.createFreezePane(0, 1); // freeze header row

        // Data rows
        int lineNum = 1;
        for (RiskFlag r : risks) {
            Row dataRow = sheet.createRow(row++);
            dataRow.setHeightInPoints(45);

            CellStyle sevStyle = switch (r.getSeverity()) {
                case HIGH   -> s.sevHigh;
                case MEDIUM -> s.sevMed;
                case LOW    -> s.sevLow;
            };

            setCell(dataRow, 0,  String.valueOf(lineNum++),                                    s.data);
            setCell(dataRow, 1,  String.valueOf(r.getPageNumber()),                             s.data);
            setCell(dataRow, 2,  r.getSeverity().name(),                                       sevStyle);
            setCell(dataRow, 3,  legalLabel(r.getCategory()),                                  s.data);

            if (r.isRedacted()) {
                setCell(dataRow, 4,  REDACTED, s.redacted);
                setCell(dataRow, 6,  REDACTED, s.redacted);
                setCell(dataRow, 9,  REDACTED, s.redacted);
            } else {
                setCell(dataRow, 4,  r.getEntityName(),  s.data);
                setCell(dataRow, 6,  r.getSnippet(),     s.data);
                setCell(dataRow, 9,  r.getComments(),    s.data);
            }

            setCell(dataRow, 5,  statusLabel(r.getStatus()),  s.data);
            setCell(dataRow, 7,  r.getReason(),               s.data);
            setCell(dataRow, 8,  r.getRestrictions(),         s.data);
            setCell(dataRow, 10, r.getSuggestion(),           s.data);
            setCell(dataRow, 11, r.isRedacted() ? "YES" : "NO",
                    r.isRedacted() ? s.redacted : s.data);
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Tab 3 — Cleared Items Summary
    // ─────────────────────────────────────────────────────────────────────────

    private void buildClearedSummary(XSSFWorkbook wb, Styles s, List<RiskFlag> risks) {
        XSSFSheet sheet = wb.createSheet("Cleared Items Summary");

        int[] widths = {8, 10, 14, 28, 40, 28, 40, 12};
        for (int i = 0; i < widths.length; i++) sheet.setColumnWidth(i, widths[i] * 256);

        // Title
        int row = 0;
        Row titleRow = sheet.createRow(row++);
        Cell tc = titleRow.createCell(0);
        tc.setCellValue("CLEARED ITEMS — For Insurer Confirmation");
        tc.setCellStyle(s.subtitle);
        sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 7));
        row++;

        // Headers
        String[] headers = {
                "Line #", "Page", "Severity", "US Legal Exposure Category",
                "Entity / Item", "Clearance Status", "Restrictions / Conditions", "Privileged"
        };
        Row headerRow = sheet.createRow(row++);
        for (int col = 0; col < headers.length; col++) {
            Cell cell = headerRow.createCell(col);
            cell.setCellValue(headers[col]);
            cell.setCellStyle(s.colHeader);
        }
        sheet.createFreezePane(0, 3);

        // Cleared rows only
        List<RiskFlag> cleared = risks.stream()
                .filter(r -> isCleared(r.getStatus()))
                .toList();

        int lineNum = 1;
        for (RiskFlag r : cleared) {
            Row dataRow = sheet.createRow(row++);
            dataRow.setHeightInPoints(40);

            CellStyle sevStyle = switch (r.getSeverity()) {
                case HIGH   -> s.sevHigh;
                case MEDIUM -> s.sevMed;
                case LOW    -> s.sevLow;
            };

            setCell(dataRow, 0, String.valueOf(lineNum++),        s.data);
            setCell(dataRow, 1, String.valueOf(r.getPageNumber()), s.data);
            setCell(dataRow, 2, r.getSeverity().name(),            sevStyle);
            setCell(dataRow, 3, legalLabel(r.getCategory()),       s.data);
            setCell(dataRow, 4, r.isRedacted() ? REDACTED : r.getEntityName(),
                    r.isRedacted() ? s.redacted : s.data);
            setCell(dataRow, 5, statusLabel(r.getStatus()),        s.cleared);
            setCell(dataRow, 6, r.getRestrictions(),               s.data);
            setCell(dataRow, 7, r.isRedacted() ? "YES" : "NO",
                    r.isRedacted() ? s.redacted : s.data);
        }

        // "No items" message if empty
        if (cleared.isEmpty()) {
            Row emptyRow = sheet.createRow(row);
            Cell ec = emptyRow.createCell(0);
            ec.setCellValue("No items have been cleared yet.");
            sheet.addMergedRegion(new CellRangeAddress(row, row, 0, 7));
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Helpers
    // ─────────────────────────────────────────────────────────────────────────

    private boolean isCleared(ClearanceStatus status) {
        return status == ClearanceStatus.CLEARED
                || status == ClearanceStatus.NO_CLEARANCE_NECESSARY
                || status == ClearanceStatus.PERMISSIBLE
                || status == ClearanceStatus.NEGOTIATED_BY_ATTORNEY
                || status == ClearanceStatus.BRANDED_INTEGRATION;
    }

    private String legalLabel(RiskCategory category) {
        return LEGAL_EXPOSURE_LABELS.getOrDefault(category.name(),
                category.name().replace("_", " "));
    }

    private String statusLabel(ClearanceStatus status) {
        return STATUS_LABELS.getOrDefault(status.name(), status.name());
    }

    private String resolveCertifyingAttorney(Script script, Long requestingUserId) {
        if (script.getProject() == null) return "—";
        return memberRepo.findByProject(script.getProject()).stream()
                .filter(m -> "ATTORNEY".equals(m.getProjectRole().name()))
                .map(m -> "@" + m.getUser().getUsername() + " (" + m.getUser().getEmail() + ")")
                .findFirst()
                .orElse("Not yet assigned");
    }

    private void setCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private int mergedRow(XSSFSheet sheet, int rowIdx, CellStyle style, String text, int colSpan) {
        Row row = sheet.createRow(rowIdx);
        row.setHeightInPoints(24);
        Cell cell = row.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        if (colSpan > 1) sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, colSpan - 1));
        return rowIdx + 1;
    }

    private int sectionHeader(XSSFSheet sheet, int rowIdx, CellStyle style, String text, int colSpan) {
        Row row = sheet.createRow(rowIdx);
        row.setHeightInPoints(20);
        Cell cell = row.createCell(0);
        cell.setCellValue(text);
        cell.setCellStyle(style);
        if (colSpan > 1) sheet.addMergedRegion(new CellRangeAddress(rowIdx, rowIdx, 0, colSpan - 1));
        return rowIdx + 1;
    }

    private int labelValue(XSSFSheet sheet, int rowIdx,
                           CellStyle labelStyle, CellStyle valueStyle,
                           String label, String value) {
        Row row = sheet.createRow(rowIdx);
        row.setHeightInPoints(18);
        Cell lc = row.createCell(0); lc.setCellValue(label); lc.setCellStyle(labelStyle);
        Cell vc = row.createCell(1); vc.setCellValue(value != null ? value : ""); vc.setCellStyle(valueStyle);
        return rowIdx + 1;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Style definitions — all colours in one inner class for clarity
    // ─────────────────────────────────────────────────────────────────────────

    private static class Styles {
        final CellStyle title, subtitle, sectionHeader, colHeader;
        final CellStyle label, value;
        final CellStyle labelHigh, valueHigh, labelMed, valueMed;
        final CellStyle labelLow, valueLow, labelPending, valuePending;
        final CellStyle labelNotClear, valueNotClear;
        final CellStyle data, redacted, cleared;
        final CellStyle sevHigh, sevMed, sevLow;
        final CellStyle certText, legalNote;

        Styles(XSSFWorkbook wb) {
            title          = build(wb, rgb(6,95,70),    white(),     16, true,  false);
            subtitle       = build(wb, rgb(15,23,42),   white(),     11, false, true);
            sectionHeader  = build(wb, rgb(31,41,55),   rgb(209,250,229), 10, true,  false);
            colHeader      = build(wb, rgb(15,23,42),   white(),     10, true,  false);
            label          = build(wb, rgb(249,250,251),rgb(55,65,81), 10, true,  false);
            value          = build(wb, white(),         rgb(15,23,42), 10, false, false);
            labelHigh      = build(wb, rgb(254,226,226),rgb(185,28,28), 10, true,  false);
            valueHigh      = build(wb, rgb(254,226,226),rgb(185,28,28), 10, false, false);
            labelMed       = build(wb, rgb(255,251,235),rgb(146,64,14), 10, true,  false);
            valueMed       = build(wb, rgb(255,251,235),rgb(146,64,14), 10, false, false);
            labelLow       = build(wb, rgb(240,253,244),rgb(21,128,61), 10, true,  false);
            valueLow       = build(wb, rgb(240,253,244),rgb(21,128,61), 10, false, false);
            labelPending   = build(wb, rgb(254,243,199),rgb(120,53,15), 10, true,  false);
            valuePending   = build(wb, rgb(254,243,199),rgb(120,53,15), 10, false, false);
            labelNotClear  = build(wb, rgb(254,226,226),rgb(127,29,29), 10, true,  false);
            valueNotClear  = build(wb, rgb(254,226,226),rgb(127,29,29), 10, false, false);
            data           = buildData(wb);
            redacted       = build(wb, rgb(30,30,30),  rgb(255,80,80), 9, true, false);
            cleared        = build(wb, rgb(240,253,244),rgb(21,128,61), 10, false, false);
            sevHigh        = build(wb, rgb(220,38,38), white(),     9, true, false);
            sevMed         = build(wb, rgb(217,119,6), white(),     9, true, false);
            sevLow         = build(wb, rgb(22,163,74), white(),     9, true, false);
            certText       = buildWrap(wb, rgb(238,253,244), rgb(20,83,45), 9, false, true);
            legalNote      = buildWrap(wb, rgb(249,250,251), rgb(55,65,81), 9, false, true);
        }

        private XSSFColor rgb(int r, int g, int b) {
            return new XSSFColor(new byte[]{(byte)r,(byte)g,(byte)b}, null);
        }
        private XSSFColor white() { return rgb(255,255,255); }

        private CellStyle build(XSSFWorkbook wb, XSSFColor bg, XSSFColor fg,
                                int pts, boolean bold, boolean italic) {
            XSSFCellStyle cs = wb.createCellStyle();
            cs.setFillForegroundColor(bg);
            cs.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            cs.setAlignment(HorizontalAlignment.LEFT);
            cs.setVerticalAlignment(VerticalAlignment.CENTER);
            cs.setBorderBottom(BorderStyle.HAIR);
            cs.setBorderRight(BorderStyle.HAIR);
            XSSFFont font = wb.createFont();
            font.setColor(fg); font.setFontHeightInPoints((short)pts);
            font.setBold(bold); font.setItalic(italic); font.setFontName("Calibri");
            cs.setFont(font);
            return cs;
        }

        private CellStyle buildData(XSSFWorkbook wb) {
            XSSFCellStyle cs = wb.createCellStyle();
            cs.setWrapText(true);
            cs.setVerticalAlignment(VerticalAlignment.TOP);
            cs.setBorderBottom(BorderStyle.HAIR);
            cs.setBorderRight(BorderStyle.HAIR);
            XSSFFont font = wb.createFont();
            font.setFontHeightInPoints((short)9); font.setFontName("Calibri");
            cs.setFont(font);
            return cs;
        }

        private CellStyle buildWrap(XSSFWorkbook wb, XSSFColor bg, XSSFColor fg,
                                    int pts, boolean bold, boolean italic) {
            CellStyle cs = build(wb, bg, fg, pts, bold, italic);
            cs.setWrapText(true);
            return cs;
        }
    }
}
