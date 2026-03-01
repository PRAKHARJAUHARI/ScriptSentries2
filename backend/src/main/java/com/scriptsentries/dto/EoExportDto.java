package com.scriptsentries.dto;

import lombok.*;
import java.time.LocalDateTime;
import java.util.List;

/**
 * E&O (Errors & Omissions) Insurance Export DTO
 *
 * Structured to match the information fields required by US E&O carriers
 * (e.g., DeWitt Stern, Front Row, Media/Entertainment Group).
 *
 * Standard E&O submission requirements for US productions:
 *   - Production identification (title, studio, MPAA rating if known)
 *   - Risk categorization by US legal exposure type
 *   - Clearance status with attorney certification
 *   - Restrictions and licensing conditions
 *   - Redaction status for privileged items
 */
public class EoExportDto {

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class EoExportRequest {
        /** Script ID to export */
        private Long scriptId;
        /**
         * If true, include only CLEARED / PERMISSIBLE / NO_CLEARANCE_NECESSARY risks.
         * If false (default), include all risks (insurer sees everything).
         */
        private boolean clearedOnlyMode;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Production Cover Sheet
    // ─────────────────────────────────────────────────────────────────────────

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class ProductionInfo {
        // Core identifiers
        private String productionTitle;
        private String studioName;
        private String director;
        private String producer;
        private String genre;
        private String logline;
        private String expectedRelease;
        private String imdbLink;
        private String contactEmail;
        private String contactPhone;

        // Script version info
        private String scriptFilename;
        private String versionLabel;
        private String analysisDate;
        private int    totalPages;

        // Summary counts
        private int    totalRisksIdentified;
        private int    highSeverityCount;
        private int    mediumSeverityCount;
        private int    lowSeverityCount;
        private int    clearedCount;
        private int    pendingCount;
        private int    notClearCount;
        private int    redactedCount;

        // Attorney certification
        private String certifyingAttorney;
        private String certificationDate;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Individual Risk Entry (one row per risk in the E&O log)
    // ─────────────────────────────────────────────────────────────────────────

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class EoRiskEntry {
        private int    lineNumber;         // Sequential for insurer reference
        private int    pageNumber;
        private String severity;           // HIGH / MEDIUM / LOW
        private String riskCategory;       // Maps to US legal category (see below)
        private String legalExposureType;  // Derived human-readable US legal label
        private String entityName;         // [REDACTED] if isRedacted
        private String scriptExcerpt;      // [REDACTED] if isRedacted
        private String legalBasis;         // reason field
        private String clearanceStatus;    // Human-readable status
        private String restrictions;       // Usage restrictions / conditions
        private String attorneyNotes;      // comments field — [REDACTED] if isRedacted
        private boolean isRedacted;
    }

    // ─────────────────────────────────────────────────────────────────────────
    // Full export package
    // ─────────────────────────────────────────────────────────────────────────

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class EoExportPackage {
        private ProductionInfo    productionInfo;
        private List<EoRiskEntry> riskLog;
    }
}
