package com.scriptsentries.dto;

import com.scriptsentries.model.Script;
import lombok.Builder;
import lombok.Data;

import java.time.LocalDateTime;
import java.util.List;

@Data
@Builder
public class ScriptResponse {
    private Long id;
    private String filename;
    private int totalPages;
    private int riskCount;
    private Script.ScriptStatus status;
    private LocalDateTime uploadedAt;
    private List<RiskFlagResponse> risks;

    public static ScriptResponse from(Script s, List<RiskFlagResponse> risks) {
        return ScriptResponse.builder()
                .id(s.getId())
                .filename(s.getFilename())
                .totalPages(s.getTotalPages())
                .riskCount(s.getRiskCount())
                .status(s.getStatus())
                .uploadedAt(s.getUploadedAt())
                .risks(risks)
                .build();
    }

    public static ScriptResponse summary(Script s) {
        return from(s, null);
    }
}
