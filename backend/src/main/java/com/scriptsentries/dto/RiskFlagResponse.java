package com.scriptsentries.dto;

import com.scriptsentries.model.RiskFlag;
import com.scriptsentries.model.enums.*;
import lombok.Builder;
import lombok.Data;

import java.time.LocalDateTime;

@Data
@Builder
public class RiskFlagResponse {
    private Long id;
    private RiskCategory category;
    private RiskSubCategory subCategory;
    private RiskSeverity severity;
    private ClearanceStatus status;
    private String entityName;
    private String snippet;
    private String reason;
    private String suggestion;
    private String comments;
    private String restrictions;
    private int pageNumber;
    private String episodeNumber;
    private String sceneNumber;
    private boolean isRedacted;
    private LocalDateTime createdAt;

    public static RiskFlagResponse from(RiskFlag f) {
        return RiskFlagResponse.builder()
                .id(f.getId())
                .category(f.getCategory())
                .subCategory(f.getSubCategory())
                .severity(f.getSeverity())
                .status(f.getStatus())
                .entityName(f.getEntityName())
                .snippet(f.getSnippet())
                .reason(f.getReason())
                .suggestion(f.getSuggestion())
                .comments(f.getComments())
                .restrictions(f.getRestrictions())
                .pageNumber(f.getPageNumber())
                .episodeNumber(f.getEpisodeNumber())
                .sceneNumber(f.getSceneNumber())
                .isRedacted(f.isRedacted())
                .createdAt(f.getCreatedAt())
                .build();
    }
}
