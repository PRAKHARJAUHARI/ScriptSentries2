package com.scriptsentries.model;

import jakarta.persistence.*;
import lombok.*;
import org.hibernate.annotations.CreationTimestamp;
import org.hibernate.annotations.Formula; // Import this

import java.time.LocalDateTime;
import java.util.List;

@Entity
@Table(name = "scripts")
@Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
public class Script {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(nullable = false)
    private String filename;

    /** * ADD THIS FIELD:
     * It allows the Repository to find 'Active' scripts.
     * The @Formula derives the value: if deleted_at is null, active is true.
     */
    @Formula("deleted_at IS NULL")
    private boolean active;

    @Column(name = "total_pages")
    private int totalPages;

    @Column(name = "risk_count")
    private int riskCount;

    @Enumerated(EnumType.STRING)
    @Column(nullable = false)
    @Builder.Default
    private ScriptStatus status = ScriptStatus.PROCESSING;

    @CreationTimestamp
    @Column(name = "uploaded_at", updatable = false)
    private LocalDateTime uploadedAt;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "project_id")
    private Project project;

    @Column(name = "version_name", length = 100)
    private String versionName;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "uploaded_by")
    private User uploadedBy;

    @Column(name = "deleted_at")
    private LocalDateTime deletedAt;

    @OneToMany(mappedBy = "script", cascade = CascadeType.ALL, fetch = FetchType.LAZY)
    private List<RiskFlag> risks;

    public boolean isDeleted() { return deletedAt != null; }

    public enum ScriptStatus {
        PROCESSING, COMPLETE, FAILED
    }
}