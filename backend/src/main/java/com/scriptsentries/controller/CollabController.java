package com.scriptsentries.controller;

import com.scriptsentries.dto.CollabDto;
import com.scriptsentries.repository.UserRepository;
import com.scriptsentries.repository.ProjectRepository;
import com.scriptsentries.repository.ProjectMemberRepository;
import com.scriptsentries.service.CommentService;
import com.scriptsentries.model.Project;
import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;

@RestController
@RequestMapping("/api/collab")
@RequiredArgsConstructor
@CrossOrigin(origins = {"http://localhost:5173", "http://localhost:3000"})
public class CollabController {

    private final CommentService commentService;
    private final UserRepository userRepository;
    private final ProjectRepository projectRepository;
    private final ProjectMemberRepository projectMemberRepository;

    // POST /api/collab/comments — save comment + trigger @mention notifications
    @PostMapping("/comments")
    public ResponseEntity<CollabDto.CommentResponse> addComment(
            @Valid @RequestBody CollabDto.CommentRequest request) {
        return ResponseEntity.status(HttpStatus.CREATED)
                .body(commentService.addComment(request));
    }

    // GET /api/collab/comments/{riskFlagId} — get all comments for a risk
    @GetMapping("/comments/{riskFlagId}")
    public ResponseEntity<List<CollabDto.CommentResponse>> getComments(
            @PathVariable Long riskFlagId) {
        return ResponseEntity.ok(commentService.getCommentsForRisk(riskFlagId));
    }

    // GET /api/collab/notifications/{userId} — get all notifications for user
    @GetMapping("/notifications/{userId}")
    public ResponseEntity<List<CollabDto.NotificationResponse>> getNotifications(
            @PathVariable Long userId) {
        return ResponseEntity.ok(commentService.getNotifications(userId));
    }

    // GET /api/collab/notifications/{userId}/unread-count
    @GetMapping("/notifications/{userId}/unread-count")
    public ResponseEntity<Long> getUnreadCount(@PathVariable Long userId) {
        return ResponseEntity.ok(commentService.getUnreadCount(userId));
    }

    // POST /api/collab/notifications/{userId}/mark-read — mark all as read
    @PostMapping("/notifications/{userId}/mark-read")
    public ResponseEntity<Void> markAllRead(@PathVariable Long userId) {
        commentService.markAllRead(userId);
        return ResponseEntity.ok().build();
    }

    // GET /api/collab/users/search?q=steve — autocomplete for @mentions
    // Optional: projectId to filter to project members only
    @GetMapping("/users/search")
    public ResponseEntity<List<CollabDto.UserSummary>> searchUsers(
            @RequestParam String q,
            @RequestParam(required = false) Long projectId) {
        if (q == null || q.isBlank() || q.length() < 1) {
            return ResponseEntity.ok(List.of());
        }
        
        List<CollabDto.UserSummary> results = userRepository
                .searchByUsername(q)
                .stream()
                .limit(8)
                .map(CollabDto.UserSummary::from)
                .toList();
        
        // If projectId provided, filter to only project members
        if (projectId != null) {
            Project project = projectRepository.findById(projectId).orElse(null);
            if (project != null) {
                results = results.stream()
                        .filter(user -> projectMemberRepository.existsByProjectAndUser(project, 
                                userRepository.findById((long) user.getId()).orElse(null)))
                        .toList();
            }
        }
        
        return ResponseEntity.ok(results);
    }
}
