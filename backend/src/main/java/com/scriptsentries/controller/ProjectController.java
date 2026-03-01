package com.scriptsentries.controller;

import com.scriptsentries.dto.ProjectDto;
import com.scriptsentries.service.ProjectService;
import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.util.List;
import java.util.Map;

@RestController
@RequestMapping("/api/projects")
@RequiredArgsConstructor
@CrossOrigin(origins = {"http://localhost:5173", "http://localhost:3000"})
public class ProjectController {

    private final ProjectService projectService;

    // ── CRUD ─────────────────────────────────────────────────────────────────

    @PostMapping
    public ResponseEntity<ProjectDto.ProjectResponse> create(
            @Valid @RequestBody ProjectDto.CreateProjectRequest req,
            @RequestParam Long userId) {
        return ResponseEntity.status(HttpStatus.CREATED).body(projectService.createProject(req, userId));
    }

    @GetMapping
    public ResponseEntity<List<ProjectDto.ProjectResponse>> list(@RequestParam Long userId) {
        return ResponseEntity.ok(projectService.getProjectsForUser(userId));
    }

    @GetMapping("/{id}")
    public ResponseEntity<ProjectDto.ProjectResponse> get(@PathVariable Long id) {
        return ResponseEntity.ok(projectService.getProject(id));
    }

    /** Update project details + production info. ATTORNEY or ANALYST only. */
    @PatchMapping("/{id}")
    public ResponseEntity<ProjectDto.ProjectResponse> update(
            @PathVariable Long id,
            @RequestBody ProjectDto.UpdateProjectRequest req,
            @RequestParam Long userId) {
        return ResponseEntity.ok(projectService.updateProject(id, req, userId));
    }

    /** Soft-delete a project. ATTORNEY only + must be a project member. */
    @DeleteMapping("/{id}")
    public ResponseEntity<ProjectDto.DeleteResponse> deleteProject(
            @PathVariable Long id,
            @RequestParam Long userId) {
        return ResponseEntity.ok(projectService.deleteProject(id, userId));
    }

    // ── TIMELINE ─────────────────────────────────────────────────────────────

    @GetMapping("/{id}/timeline")
    public ResponseEntity<ProjectDto.ProjectTimeline> timeline(@PathVariable Long id) {
        return ResponseEntity.ok(projectService.getTimeline(id));
    }

    // ── MEMBERS ──────────────────────────────────────────────────────────────

    @PostMapping("/{id}/members")
    public ResponseEntity<ProjectDto.MemberResponse> addMember(
            @PathVariable Long id,
            @Valid @RequestBody ProjectDto.MemberInvite invite,
            @RequestParam Long requestingUserId) {
        return ResponseEntity.status(HttpStatus.CREATED)
                .body(projectService.addMember(id, invite, requestingUserId));
    }

    /** Remove a member from a project. ATTORNEY or ANALYST only. Cannot remove creator. */
    @DeleteMapping("/{id}/members/{targetUserId}")
    public ResponseEntity<?> removeMember(
            @PathVariable Long id,
            @PathVariable Long targetUserId,
            @RequestParam Long requestingUserId) {
        try {
            projectService.removeMember(id, targetUserId, requestingUserId);
            return ResponseEntity.ok(Map.of("message", "Member removed"));
        } catch (RuntimeException e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }

    // ── SCRIPTS ──────────────────────────────────────────────────────────────

    /**
     * Assign script to project with version label.
     * Enforces: uploader must be a project member with upload permission (not VIEWER).
     */
    @PostMapping("/scripts/{scriptId}/assign")
    public ResponseEntity<?> assign(
            @PathVariable Long scriptId,
            @RequestParam Long projectId,
            @RequestParam(required = false) String versionName,
            @RequestParam Long userId) {
        try {
            projectService.assignScriptToProject(scriptId, projectId, versionName, userId);
            return ResponseEntity.ok(Map.of("message", "Script assigned"));
        } catch (RuntimeException e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }

    /** Soft-delete a script. ATTORNEY + ANALYST only, must be project member. */
    @DeleteMapping("/scripts/{scriptId}")
    public ResponseEntity<ProjectDto.DeleteResponse> deleteScript(
            @PathVariable Long scriptId,
            @RequestParam Long userId) {
        return ResponseEntity.ok(projectService.deleteScript(scriptId, userId));
    }

    @PatchMapping("/scripts/{scriptId}/rename")
    public ResponseEntity<?> rename(
            @PathVariable Long scriptId,
            @Valid @RequestBody ProjectDto.RenameVersionRequest req,
            @RequestParam Long userId) {
        try {
            projectService.renameVersion(scriptId, req.getVersionName(), userId);
            return ResponseEntity.ok(Map.of("message", "Version renamed"));
        } catch (RuntimeException e) {
            return ResponseEntity.badRequest().body(Map.of("error", e.getMessage()));
        }
    }
}
