package com.scriptsentries.model;

import com.scriptsentries.model.enums.*;
import jakarta.persistence.*;
import lombok.*;
import org.hibernate.annotations.CreationTimestamp;

import java.time.LocalDateTime;

@Entity
@Table(name = "risk_flags")
@Getter
@Setter
@NoArgsConstructor
@AllArgsConstructor
@Builder
public class RiskFlag {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Enumerated(EnumType.STRING)
    @Column(nullable = true)
    private RiskCategory category;

    @Enumerated(EnumType.STRING)
    @Column(nullable = true)
    private RiskSubCategory subCategory;

    @Enumerated(EnumType.STRING)
    @Column(nullable = true)
    private RiskSeverity severity;

    @Enumerated(EnumType.STRING)
    @Column(nullable = true)
    @Builder.Default
    private ClearanceStatus status = ClearanceStatus.PENDING;

    @Column(nullable = true)
    private String entityName;

    @Column(columnDefinition = "TEXT")
    private String snippet;

    @Column(columnDefinition = "TEXT", nullable = true)
    private String reason;

    @Column(columnDefinition = "TEXT")
    private String suggestion;

    @Column(columnDefinition = "TEXT")
    private String comments;

    @Column(columnDefinition = "TEXT")
    private String restrictions;

    @Column(nullable = true)
    private int pageNumber;

    @Column(nullable = true)
    private String episodeNumber;

    @Column(nullable = true)
    private String sceneNumber;

    /**
     * Security field: when true, sensitive columns are replaced with [REDACTED]
     * in any exported document. Controlled exclusively by attorneys.
     */
    @Column(nullable = true)
    @Builder.Default
    private boolean isRedacted = true;

    @CreationTimestamp
    @Column(updatable = true)
    private LocalDateTime createdAt;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "script_id", nullable = true)
    private Script script;
}
