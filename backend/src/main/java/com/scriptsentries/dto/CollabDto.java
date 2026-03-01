package com.scriptsentries.dto;

import com.scriptsentries.model.Comment;
import com.scriptsentries.model.Notification;
import com.scriptsentries.model.User;
import jakarta.validation.constraints.NotBlank;
import jakarta.validation.constraints.NotNull;
import lombok.Builder;
import lombok.Data;
import lombok.AllArgsConstructor;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;

public class CollabDto {

    @Data
    public static class CommentRequest {
        @NotBlank(message = "Comment text is required")
        private String text;

        @NotNull(message = "Risk flag ID is required")
        private Long riskFlagId;

        @NotNull(message = "Author ID is required")
        private Long authorId;
    }

    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class CommentResponse {
        private Long id;
        private String text;
        private UserSummary author;
        private Long riskFlagId;
        private LocalDateTime createdAt;

        public static CommentResponse from(Comment c) {
            return CommentResponse.builder()
                    .id(c.getId())
                    .text(c.getText())
                    .author(UserSummary.from(c.getAuthor()))
                    .riskFlagId(c.getRiskFlag().getId())
                    .createdAt(c.getCreatedAt())
                    .build();
        }
    }

    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class NotificationResponse {
        private Long id;
        private String message;
        private boolean isRead;
        private Long riskFlagId;
        private LocalDateTime createdAt;

        public static NotificationResponse from(Notification n) {
            return NotificationResponse.builder()
                    .id(n.getId())
                    .message(n.getMessage())
                    .isRead(n.isRead())
                    .riskFlagId(n.getRiskFlag() != null ? n.getRiskFlag().getId() : null)
                    .createdAt(n.getCreatedAt())
                    .build();
        }
    }

    @Data
    @Builder
    @NoArgsConstructor
    @AllArgsConstructor
    public static class UserSummary {
        private Long id;
        private String username;
        private String email;
        private User.UserRole role;

        public static UserSummary from(User u) {
            return UserSummary.builder()
                    .id(u.getId())
                    .username(u.getUsername())
                    .email(u.getEmail())
                    .role(u.getRole())
                    .build();
        }
    }
}
