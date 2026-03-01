package com.scriptsentries.dto;

import com.scriptsentries.model.enums.ClearanceStatus;
import lombok.Data;

@Data
public class RiskUpdateRequest {
    private ClearanceStatus status;
    private String comments;
    private String restrictions;
    private Boolean isRedacted;
}
