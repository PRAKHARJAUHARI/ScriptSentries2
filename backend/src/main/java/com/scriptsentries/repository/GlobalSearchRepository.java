package com.scriptsentries.repository;

import com.scriptsentries.model.RiskFlag;
import com.scriptsentries.model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface GlobalSearchRepository extends JpaRepository<RiskFlag, Long> {

    /**
     * Full cross-project search.
     *
     * Finds all non-deleted risk flags whose entityName contains the query string
     * (case-insensitive), across every active (non-deleted) script in every
     * active (non-deleted) project that the requesting user is a member of.
     *
     * This enforces the project membership boundary — users only see results
     * from projects they belong to.
     *
     * Results ordered: project name → script uploadedAt DESC → page number
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
     * Scoped version: search within a single project only.
     * Used when the user has already selected a project and wants to search
     * across its script versions.
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
