package com.scriptsentries.dto;

import com.scriptsentries.model.enums.*;
import lombok.Data;
import lombok.NoArgsConstructor;
import lombok.AllArgsConstructor;

import java.util.List;

/**
 * Structured output from the AI model, parsed via Spring AI BeanOutputParser.
 */
@Data
@NoArgsConstructor
@AllArgsConstructor
public class AiPageAnalysisResult {

    private List<AiRiskItem> risks;

    @Data
    @NoArgsConstructor
    @AllArgsConstructor
    public static class AiRiskItem {
        private String category;        // maps to RiskCategory
        private String subCategory;     // maps to RiskSubCategory
        private String severity;        // HIGH | MEDIUM | LOW
        private String status;          // ClearanceStatus suggestion
        private String entityName;
        private String snippet;
        private String reason;
        private String suggestion;
        private String episodeNumber;   // Extract from script context
        private String sceneNumber;     // Extract from script context
    }
}
