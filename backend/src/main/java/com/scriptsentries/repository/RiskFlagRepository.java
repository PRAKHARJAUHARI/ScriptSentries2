package com.scriptsentries.repository;

import com.scriptsentries.model.RiskFlag;
import com.scriptsentries.model.Script;
import com.scriptsentries.model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface RiskFlagRepository extends JpaRepository<RiskFlag, Long> {


    List<RiskFlag> findByScriptOrderBySeverityAsc(Script script);
    // ── Existing ──────────────────────────────────────────────────────────────

    List<RiskFlag> findByScriptOrderByPageNumberAscSeverityDesc(Script script);

    // ── Global Search ─────────────────────────────────────────────────────────

    /**
     * Cross-project search: finds risk flags whose entityName contains the query,
     * scoped to projects the requesting user is a member of.
     */
    @Query("""
        SELECT r FROM RiskFlag r
        JOIN r.script s
        JOIN s.project p
        JOIN p.members m
        WHERE m.user = :user
          AND p.deletedAt IS NULL
          AND s.deletedAt IS NULL
          AND LOWER(r.entityName) LIKE LOWER(CONCAT('%', :query, '%'))
        ORDER BY p.name ASC, s.uploadedAt DESC, r.pageNumber ASC
        """)
    List<RiskFlag> searchByEntityNameAcrossProjects(
            @Param("user") User user,
            @Param("query") String query
    );

    /**
     * Scoped search within a single project across all its script versions.
     */
    @Query("""
        SELECT r FROM RiskFlag r
        JOIN r.script s
        JOIN s.project p
        WHERE p.id = :projectId
          AND s.deletedAt IS NULL
          AND LOWER(r.entityName) LIKE LOWER(CONCAT('%', :query, '%'))
        ORDER BY s.uploadedAt DESC, r.pageNumber ASC
        """)
    List<RiskFlag> searchByEntityNameInProject(
            @Param("projectId") Long projectId,
            @Param("query") String query
    );
}