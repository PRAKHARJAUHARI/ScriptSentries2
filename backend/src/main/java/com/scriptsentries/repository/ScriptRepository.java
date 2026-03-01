package com.scriptsentries.repository;

import com.scriptsentries.model.Project;
import com.scriptsentries.model.Script;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface ScriptRepository extends JpaRepository<Script, Long> {

    // ── Soft-delete aware ────────────────────────────────────────────────────

    /** Active (non-deleted) scripts for a project, newest first, with risks fetched. */
    @Query("SELECT s FROM Script s LEFT JOIN FETCH s.risks " +
            "WHERE s.project = :project AND s.deletedAt IS NULL " +
            "ORDER BY s.uploadedAt DESC")
    List<Script> findActiveByProjectOrderByUploadedAtDesc(@Param("project") Project project);

    /** All scripts including deleted (for audit/admin views). */
    @Query("SELECT s FROM Script s LEFT JOIN FETCH s.risks " +
            "WHERE s.project = :project ORDER BY s.uploadedAt DESC")
    List<Script> findAllByProjectOrderByUploadedAtDesc(@Param("project") Project project);

    /** Count active (non-deleted) versions in a project. */
    long countByProjectAndDeletedAtIsNull(Project project);

    /** Active scripts across all projects (homepage Recent Activity). */
    @Query("SELECT s FROM Script s WHERE s.deletedAt IS NULL ORDER BY s.uploadedAt DESC")
    List<Script> findAllActiveOrderByUploadedAtDesc();

    /**
     * Legacy method name kept for backward compatibility.
     * Any existing code calling findAllByOrderByUploadedAtDesc() will still compile.
     * Returns ALL scripts including soft-deleted — prefer findAllActiveOrderByUploadedAtDesc().
     */
    List<Script> findAllByOrderByUploadedAtDesc();
}