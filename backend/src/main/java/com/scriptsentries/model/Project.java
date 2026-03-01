package com.scriptsentries.model;

import jakarta.persistence.*;
import lombok.*;
import org.hibernate.annotations.CreationTimestamp;

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;

/**
 * V3 additions:
 *  - Production detail fields (director, producer, genre, logline, etc.)
 *  - deleted_at for soft-delete (null = active)
 */
@Entity
@Table(name = "projects")
@Getter @Setter @NoArgsConstructor @AllArgsConstructor @Builder
public class Project {

    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(nullable = false)
    private String name;

    @Column(name = "studio_name")
    private String studioName;

    // ── Production detail fields (V3) ─────────────────────────────────────────
    private String director;
    private String producer;

    @Column(name = "production_email")
    private String productionEmail;

    @Column(name = "production_phone")
    private String productionPhone;

    private String genre;

    @Column(columnDefinition = "TEXT")
    private String logline;

    @Column(name = "expected_release")
    private String expectedRelease;

    @Column(name = "imdb_link")
    private String imdbLink;

    @Column(columnDefinition = "TEXT")
    private String notes;
    // ─────────────────────────────────────────────────────────────────────────

    @CreationTimestamp
    @Column(name = "created_at", updatable = false)
    private LocalDateTime createdAt;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "created_by")
    private User createdBy;

    /** Soft-delete timestamp. NULL = active project. */
    @Column(name = "deleted_at")
    private LocalDateTime deletedAt;

    @OneToMany(mappedBy = "project", cascade = CascadeType.ALL, fetch = FetchType.LAZY)
    @Builder.Default
    private List<Script> scripts = new ArrayList<>();

    @OneToMany(mappedBy = "project", cascade = CascadeType.ALL, fetch = FetchType.LAZY)
    @Builder.Default
    private List<ProjectMember> members = new ArrayList<>();

    public boolean isDeleted() { return deletedAt != null; }
}
