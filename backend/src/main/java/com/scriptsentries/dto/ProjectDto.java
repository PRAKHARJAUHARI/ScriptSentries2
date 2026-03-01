package com.scriptsentries.dto;

import com.scriptsentries.model.*;
import com.scriptsentries.model.ProjectRole;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotNull;
import lombok.*;

import java.time.LocalDateTime;
import java.util.List;

public class ProjectDto {

    // ── Requests ──────────────────────────────────────────────────────────────

    @Data
    public static class CreateProjectRequest {
        @NotBlank(message = "Project name required")
        private String name;
        private String studioName;
        private String director;
        private String producer;
        private String productionEmail;
        private String productionPhone;
        private String genre;
        private String logline;
        private String expectedRelease;
        private String imdbLink;
        private String notes;
        private List<MemberInvite> members;
    }

    /** PATCH /api/projects/{id} — update production details + manage members */
    @Data
    public static class UpdateProjectRequest {
        private String name;
        private String studioName;
        private String director;
        private String producer;
        private String productionEmail;
        private String productionPhone;
        private String genre;
        private String logline;
        private String expectedRelease;
        private String imdbLink;
        private String notes;
    }

    @Data
    public static class MemberInvite {
        @NotNull private Long userId;
        @NotNull private ProjectRole projectRole;
    }

    @Data
    public static class RenameVersionRequest {
        @NotBlank private String versionName;
    }

    // ── Responses ─────────────────────────────────────────────────────────────

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class ProjectResponse {
        private Long                 id;
        private String               name;
        private String               studioName;
        // Production details
        private String               director;
        private String               producer;
        private String               productionEmail;
        private String               productionPhone;
        private String               genre;
        private String               logline;
        private String               expectedRelease;
        private String               imdbLink;
        private String               notes;
        private LocalDateTime        createdAt;
        private UserSummary          createdBy;
        private List<MemberResponse> members;
        private int                  totalScripts;
        private int                  totalRisks;

        public static ProjectResponse from(Project p, List<ProjectMember> members) {
            return ProjectResponse.builder()
                    .id(p.getId()).name(p.getName()).studioName(p.getStudioName())
                    .director(p.getDirector()).producer(p.getProducer())
                    .productionEmail(p.getProductionEmail()).productionPhone(p.getProductionPhone())
                    .genre(p.getGenre()).logline(p.getLogline())
                    .expectedRelease(p.getExpectedRelease()).imdbLink(p.getImdbLink())
                    .notes(p.getNotes())
                    .createdAt(p.getCreatedAt())
                    .createdBy(p.getCreatedBy() != null ? UserSummary.from(p.getCreatedBy()) : null)
                    .members(members.stream().map(MemberResponse::from).toList())
                    .totalScripts(p.getScripts() != null
                            ? (int) p.getScripts().stream().filter(s -> !s.isDeleted()).count() : 0)
                    .totalRisks(p.getScripts() != null
                            ? p.getScripts().stream().filter(s -> !s.isDeleted())
                            .mapToInt(Script::getRiskCount).sum() : 0)
                    .build();
        }
    }

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class MemberResponse {
        private Long          id;
        private UserSummary   user;
        private ProjectRole   projectRole;
        private LocalDateTime joinedAt;

        public static MemberResponse from(ProjectMember pm) {
            return MemberResponse.builder()
                    .id(pm.getId()).user(UserSummary.from(pm.getUser()))
                    .projectRole(pm.getProjectRole()).joinedAt(pm.getJoinedAt())
                    .build();
        }
    }

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class UserSummary {
        private Long   id;
        private String username;
        private String email;
        private String role;

        public static UserSummary from(User u) {
            return UserSummary.builder()
                    .id(u.getId()).username(u.getUsername()).email(u.getEmail())
                    .role(u.getRole().name()).build();
        }
    }

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class DeleteResponse {
        private String message;
        private Long   id;
    }

    // ── Timeline ──────────────────────────────────────────────────────────────

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class TimelineEntry {
        private Long          scriptId;
        private String        filename;
        private String        versionName;
        private LocalDateTime uploadedAt;
        private LocalDateTime deletedAt;
        private String        status;
        private int           totalPages;
        private int           highCount;
        private int           mediumCount;
        private int           lowCount;
        private int           totalRisks;
        private UserSummary   uploadedBy;

        public static TimelineEntry from(Script s) {
            int high = 0, med = 0, low = 0;
            if (s.getRisks() != null) {
                for (var r : s.getRisks()) {
                    switch (r.getSeverity()) {
                        case HIGH   -> high++;
                        case MEDIUM -> med++;
                        case LOW    -> low++;
                    }
                }
            }
            return TimelineEntry.builder()
                    .scriptId(s.getId()).filename(s.getFilename())
                    .versionName(s.getVersionName() != null ? s.getVersionName() : "Unnamed Version")
                    .uploadedAt(s.getUploadedAt()).deletedAt(s.getDeletedAt())
                    .status(s.getStatus().name()).totalPages(s.getTotalPages())
                    .highCount(high).mediumCount(med).lowCount(low)
                    .totalRisks(s.getRiskCount())
                    .uploadedBy(s.getUploadedBy() != null ? UserSummary.from(s.getUploadedBy()) : null)
                    .build();
        }
    }

    @Data @Builder @NoArgsConstructor @AllArgsConstructor
    public static class ProjectTimeline {
        private Long                projectId;
        private String              projectName;
        private String              studioName;
        private List<TimelineEntry> versions;
        private int                 totalVersions;
        private int                 totalHighRisks;
    }
}
