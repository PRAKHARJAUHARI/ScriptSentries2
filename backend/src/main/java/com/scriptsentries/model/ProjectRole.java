package com.scriptsentries.model;

/**
 * Unified role enum used for both User.role (global) and ProjectMember.projectRole (per-project).
 *
 * Permission matrix (highest → lowest):
 * ┌─────────────────────────┬───┬────────────────────────────────────────────────────────┐
 * │ Role                    │ # │ Permissions                                            │
 * ├─────────────────────────┼───┼────────────────────────────────────────────────────────┤
 * │ ATTORNEY                │ 5 │ Full access + finalize + delete project/scripts        │
 * │ ANALYST                 │ 4 │ Edit risks, comment, delete scripts (not projects)     │
 * │ MAIN_PRODUCTION_CONTACT │ 3 │ View + rename version label                           │
 * │ PRODUCTION_ASSISTANT    │ 2 │ View only                                              │
 * │ VIEWER                  │ 1 │ Read-only, no edits, no uploads — contact list only   │
 * └─────────────────────────┴───┴────────────────────────────────────────────────────────┘
 *
 * Who can add a VIEWER to a project: ANALYST and above.
 */
public enum ProjectRole {
    ATTORNEY,
    ANALYST,
    MAIN_PRODUCTION_CONTACT,
    PRODUCTION_ASSISTANT,
    VIEWER;

    /** Can upload scripts to a project. */
    public boolean canUpload() {
        return this == ATTORNEY || this == ANALYST || this == MAIN_PRODUCTION_CONTACT || this == PRODUCTION_ASSISTANT;
    }

    /** Can edit risk status / comments / restrictions. */
    public boolean canEdit() {
        return this == ATTORNEY || this == ANALYST;
    }

    /** Can finalize and lock a clearance report. */
    public boolean canFinalize() {
        return this == ATTORNEY;
    }

    /** Can rename a script version label. */
    public boolean canRename() {
        return this == ATTORNEY || this == MAIN_PRODUCTION_CONTACT;
    }

    /** Can delete scripts (soft-delete) within a project. */
    public boolean canDeleteScript() {
        return this == ATTORNEY || this == ANALYST;
    }

    /** Can delete the entire project. */
    public boolean canDeleteProject() {
        return this == ATTORNEY;
    }

    /** Can manage project members (add / remove / change role). */
    public boolean canManageMembers() {
        return this == ATTORNEY || this == ANALYST;
    }

    /** Can add a VIEWER to the contact list (ANALYST and above). */
    public boolean canAddViewer() {
        return this == ATTORNEY || this == ANALYST;
    }

    /** True for completely read-only roles. */
    public boolean isReadOnly() {
        return this == VIEWER || this == PRODUCTION_ASSISTANT;
    }
}
