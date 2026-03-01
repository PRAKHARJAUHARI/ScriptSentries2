package com.scriptsentries.dto;

import lombok.*;
import java.time.LocalDateTime;
import java.util.List;

public class GlobalSearchDto {

    // ── Request ───────────────────────────────────────────────────────────────

    /** GET /api/search?q=John+Wick&userId=7  — no request body needed */

    // ── Response ──────────────────────────────────────────────────────────────

    /**
     * One matching risk flag hit, enriched with project + script context
     * so the UI can show "Project: The Last Horizon  ›  Draft 2  ›  pg.14"
     */
    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class SearchHit {
        // Risk flag data
        private Long   riskFlagId;
        private String entityName;
        private String category;
        private String subCategory;
        private String severity;
        private String status;
        private String snippet;
        private int    pageNumber;

        // Script context
        private Long   scriptId;
        private String scriptFilename;
        private String versionName;

        // Project context
        private Long   projectId;
        private String projectName;
        private String studioName;
    }

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class SearchResponse {
        private String        query;
        private int           totalHits;
        private List<SearchHit> hits;
    }
}
