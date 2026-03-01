package com.scriptsentries.repository;

import com.scriptsentries.model.Project;
import com.scriptsentries.model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface ProjectRepository extends JpaRepository<Project, Long> {

    /** Active (non-deleted) projects the user is a member of, newest first. */
    @Query("SELECT p FROM Project p JOIN p.members m " +
            "WHERE m.user = :user AND p.deletedAt IS NULL " +
            "ORDER BY p.createdAt DESC")
    List<Project> findActiveByMember(@Param("user") User user);

    /** All projects (including deleted) the user is a member of. */
    @Query("SELECT p FROM Project p JOIN p.members m WHERE m.user = :user ORDER BY p.createdAt DESC")
    List<Project> findAllByMember(@Param("user") User user);
}
