package com.scriptsentries.repository;

import com.scriptsentries.model.Project;
import com.scriptsentries.model.ProjectMember;
import com.scriptsentries.model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface ProjectMemberRepository extends JpaRepository<ProjectMember, Long> {
    List<ProjectMember>     findByProject(Project project);
    Optional<ProjectMember> findByProjectAndUser(Project project, User user);
    boolean                 existsByProjectAndUser(Project project, User user);

    // UserRepository (keep your existing file, just add these if not already there)
    // Optional<User> findByEmail(String email);
    // Optional<User> findByUsername(String username);
    // boolean existsByEmail(String email);
    // boolean existsByUsername(String username);
    // List<User> findByUsernameContainingIgnoreCase(String query);
}
