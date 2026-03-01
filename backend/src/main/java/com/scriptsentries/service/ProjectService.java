package com.scriptsentries.service;

import com.scriptsentries.dto.ProjectDto;
import com.scriptsentries.model.*;
import com.scriptsentries.model.ProjectRole;
import com.scriptsentries.repository.*;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.time.LocalDateTime;
import java.util.List;

@Service
@RequiredArgsConstructor
@Slf4j
public class ProjectService {

    private final ProjectRepository       projectRepo;
    private final ProjectMemberRepository memberRepo;
    private final UserRepository          userRepo;
    private final ScriptRepository        scriptRepo;
    private final NotificationRepository  notificationRepo;

    // ─────────────────────────────────────────────────────────────────────────
    // CREATE PROJECT
    // ─────────────────────────────────────────────────────────────────────────

    @Transactional
    public ProjectDto.ProjectResponse createProject(ProjectDto.CreateProjectRequest req, Long creatorId) {
        User creator = findUser(creatorId);

        Project project = Project.builder()
                .name(req.getName())
                .studioName(req.getStudioName())
                .director(req.getDirector())
                .producer(req.getProducer())
                .productionEmail(req.getProductionEmail())
                .productionPhone(req.getProductionPhone())
                .genre(req.getGenre())
                .logline(req.getLogline())
                .expectedRelease(req.getExpectedRelease())
                .imdbLink(req.getImdbLink())
                .notes(req.getNotes())
                .createdBy(creator)
                .build();
        project = projectRepo.save(project);

        // Creator always gets ATTORNEY role
        memberRepo.save(ProjectMember.builder()
                .project(project).user(creator)
                .projectRole(ProjectRole.ATTORNEY).build());

        if (req.getMembers() != null) {
            final Project saved = project;
            req.getMembers().stream()
                    .filter(inv -> !inv.getUserId().equals(creatorId))
                    .forEach(inv -> userRepo.findById(inv.getUserId()).ifPresent(u -> {
                        if (!memberRepo.existsByProjectAndUser(saved, u)) {
                            memberRepo.save(ProjectMember.builder()
                                    .project(saved).user(u)
                                    .projectRole(inv.getProjectRole()).build());
                        }
                    }));
        }

        List<ProjectMember> members = memberRepo.findByProject(project);
        log.info("Project '{}' created by @{}", project.getName(), creator.getUsername());
        return ProjectDto.ProjectResponse.from(project, members);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // UPDATE PROJECT DETAILS  (ATTORNEY + ANALYST can update)
    // ─────────────────────────────────────────────────────────────────────────

    @Transactional
    public ProjectDto.ProjectResponse updateProject(Long projectId, ProjectDto.UpdateProjectRequest req, Long requestingUserId) {
        Project project  = findProject(projectId);
        User    requester = findUser(requestingUserId);
        ProjectMember pm  = requireMembership(project, requester);

        if (!pm.getProjectRole().canEdit()) {
            throw new RuntimeException("Only ATTORNEY or ANALYST can update project details");
        }

        if (req.getName()            != null) project.setName(req.getName());
        if (req.getStudioName()      != null) project.setStudioName(req.getStudioName());
        if (req.getDirector()        != null) project.setDirector(req.getDirector());
        if (req.getProducer()        != null) project.setProducer(req.getProducer());
        if (req.getProductionEmail() != null) project.setProductionEmail(req.getProductionEmail());
        if (req.getProductionPhone() != null) project.setProductionPhone(req.getProductionPhone());
        if (req.getGenre()           != null) project.setGenre(req.getGenre());
        if (req.getLogline()         != null) project.setLogline(req.getLogline());
        if (req.getExpectedRelease() != null) project.setExpectedRelease(req.getExpectedRelease());
        if (req.getImdbLink()        != null) project.setImdbLink(req.getImdbLink());
        if (req.getNotes()           != null) project.setNotes(req.getNotes());

        projectRepo.save(project);
        return ProjectDto.ProjectResponse.from(project, memberRepo.findByProject(project));
    }

    // ─────────────────────────────────────────────────────────────────────────
    // READ
    // ─────────────────────────────────────────────────────────────────────────

    public List<ProjectDto.ProjectResponse> getProjectsForUser(Long userId) {
        User user = findUser(userId);
        return projectRepo.findActiveByMember(user).stream()
                .map(p -> ProjectDto.ProjectResponse.from(p, memberRepo.findByProject(p)))
                .toList();
    }

    public ProjectDto.ProjectResponse getProject(Long projectId) {
        Project p = findProject(projectId);
        return ProjectDto.ProjectResponse.from(p, memberRepo.findByProject(p));
    }

    public ProjectDto.ProjectTimeline getTimeline(Long projectId) {
        Project project = findProject(projectId);
        List<Script> scripts = scriptRepo.findActiveByProjectOrderByUploadedAtDesc(project);
        List<ProjectDto.TimelineEntry> entries = scripts.stream()
                .map(ProjectDto.TimelineEntry::from).toList();
        int totalHigh = entries.stream().mapToInt(ProjectDto.TimelineEntry::getHighCount).sum();

        return ProjectDto.ProjectTimeline.builder()
                .projectId(project.getId()).projectName(project.getName())
                .studioName(project.getStudioName()).versions(entries)
                .totalVersions(entries.size()).totalHighRisks(totalHigh)
                .build();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // DELETE PROJECT  (ATTORNEY only, must be a member)
    // ─────────────────────────────────────────────────────────────────────────

    @Transactional
    public ProjectDto.DeleteResponse deleteProject(Long projectId, Long requestingUserId) {
        Project project  = findProject(projectId);
        User    requester = findUser(requestingUserId);

        // Must be in the project's contact list
        ProjectMember pm = requireMembership(project, requester);

        // Only ATTORNEY can delete the project
        if (!pm.getProjectRole().canDeleteProject()) {
            throw new RuntimeException("Only ATTORNEY role can delete a project");
        }

        // Soft-delete project and all its scripts
        LocalDateTime now = LocalDateTime.now();
        project.setDeletedAt(now);
        project.getScripts().forEach(s -> s.setDeletedAt(now));
        projectRepo.save(project);

        log.info("Project '{}' soft-deleted by @{}", project.getName(), requester.getUsername());
        return ProjectDto.DeleteResponse.builder()
                .message("Project deleted successfully").id(projectId).build();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // DELETE SCRIPT  (ATTORNEY + ANALYST, must be a project member)
    // ─────────────────────────────────────────────────────────────────────────

    @Transactional
    public ProjectDto.DeleteResponse deleteScript(Long scriptId, Long requestingUserId) {
        Script script    = findScript(scriptId);
        User   requester = findUser(requestingUserId);

        // Must be a member of the project this script belongs to
        if (script.getProject() == null) {
            throw new RuntimeException("Script is not associated with any project");
        }
        ProjectMember pm = requireMembership(script.getProject(), requester);

        if (!pm.getProjectRole().canDeleteScript()) {
            throw new RuntimeException("VIEWER, PRODUCTION_ASSISTANT, and MAIN_PRODUCTION_CONTACT cannot delete scripts");
        }

        script.setDeletedAt(LocalDateTime.now());
        scriptRepo.save(script);

        log.info("Script '{}' soft-deleted by @{}", script.getFilename(), requester.getUsername());
        return ProjectDto.DeleteResponse.builder()
                .message("Script deleted successfully").id(scriptId).build();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // MEMBERS  (add / remove)
    // ─────────────────────────────────────────────────────────────────────────

    @Transactional
    public ProjectDto.MemberResponse addMember(Long projectId, ProjectDto.MemberInvite invite, Long requestingUserId) {
        Project project   = findProject(projectId);
        User    requester = findUser(requestingUserId);
        ProjectMember rm  = requireMembership(project, requester);

        // ANALYST and above can add members (including VIEWERs)
        if (!rm.getProjectRole().canManageMembers()) {
            throw new RuntimeException("Only ATTORNEY or ANALYST can add members");
        }

        // Only ANALYST+ can add a VIEWER
        if (invite.getProjectRole() == ProjectRole.VIEWER && !rm.getProjectRole().canAddViewer()) {
            throw new RuntimeException("Only ANALYST or ATTORNEY can add VIEWER to a project");
        }

        User newUser = findUser(invite.getUserId());
        if (memberRepo.existsByProjectAndUser(project, newUser)) {
            throw new RuntimeException("User is already a member");
        }

        ProjectMember member = memberRepo.save(ProjectMember.builder()
                .project(project).user(newUser)
                .projectRole(invite.getProjectRole()).build());
        
        // Create notification for the new member
        String message = String.format("%s added you to project '%s'", requester.getUsername(), project.getName());
        notificationRepo.save(Notification.builder()
                .recipient(newUser)
                .message(message)
                .riskFlag(null)
                .isRead(false)
                .build());
        
        log.info("@{} added to project '{}', notification sent", newUser.getUsername(), project.getName());
        
        return ProjectDto.MemberResponse.from(member);
    }

    @Transactional
    public void removeMember(Long projectId, Long targetUserId, Long requestingUserId) {
        Project project   = findProject(projectId);
        User    requester = findUser(requestingUserId);
        ProjectMember rm  = requireMembership(project, requester);

        if (!rm.getProjectRole().canManageMembers()) {
            throw new RuntimeException("Only ATTORNEY or ANALYST can remove members");
        }

        User target = findUser(targetUserId);

        // Cannot remove the project creator
        if (project.getCreatedBy() != null && project.getCreatedBy().getId().equals(targetUserId)) {
            throw new RuntimeException("Cannot remove the project creator");
        }

        memberRepo.findByProjectAndUser(project, target).ifPresent(memberRepo::delete);
        log.info("@{} removed from project '{}'", target.getUsername(), project.getName());
    }

    // ─────────────────────────────────────────────────────────────────────────
    // SCRIPT VERSIONING
    // ─────────────────────────────────────────────────────────────────────────

    /**
     * Enforce: no upload without a project.
     * Called from ScriptAnalysisService after analysis completes.
     */
    @Transactional
    public void assignScriptToProject(Long scriptId, Long projectId, String versionName, Long uploaderId) {
        Script  script  = findScript(scriptId);
        Project project = findProject(projectId);
        User    uploader = findUser(uploaderId);

        // Verify uploader is a member with upload permission
        ProjectMember pm = requireMembership(project, uploader);
        if (!pm.getProjectRole().canUpload()) {
            throw new RuntimeException("VIEWER role cannot upload scripts");
        }

        if (versionName == null || versionName.isBlank()) {
            long count = scriptRepo.countByProjectAndDeletedAtIsNull(project);
            versionName = "Draft " + (count + 1);
        }

        script.setProject(project);
        script.setVersionName(versionName);
        script.setUploadedBy(uploader);
        scriptRepo.save(script);
    }

    @Transactional
    public void renameVersion(Long scriptId, String newName, Long requestingUserId) {
        Script script    = findScript(scriptId);
        User   requester = findUser(requestingUserId);

        if (script.getProject() != null) {
            ProjectMember m = requireMembership(script.getProject(), requester);
            if (!m.getProjectRole().canRename()) {
                throw new RuntimeException("Only ATTORNEY or MAIN_PRODUCTION_CONTACT can rename versions");
            }
        }

        script.setVersionName(newName);
        scriptRepo.save(script);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // HELPERS
    // ─────────────────────────────────────────────────────────────────────────

    private User        findUser(Long id)    { return userRepo.findById(id).orElseThrow(()    -> new RuntimeException("User not found: "    + id)); }
    private Project     findProject(Long id) { return projectRepo.findById(id).orElseThrow(() -> new RuntimeException("Project not found: " + id)); }
    private Script      findScript(Long id)  { return scriptRepo.findById(id).orElseThrow(()  -> new RuntimeException("Script not found: "  + id)); }

    private ProjectMember requireMembership(Project project, User user) {
        return memberRepo.findByProjectAndUser(project, user)
                .orElseThrow(() -> new RuntimeException(
                        "User @" + user.getUsername() + " is not a member of project '" + project.getName() + "'"));
    }
}
