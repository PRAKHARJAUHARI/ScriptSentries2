package com.scriptsentries.repository;

import com.scriptsentries.model.User;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.data.repository.query.Param;
import org.springframework.stereotype.Repository;

import java.util.List;
import java.util.Optional;

@Repository
public interface UserRepository extends JpaRepository<User, Long> {
    Optional<User> findByEmail(String email);
    Optional<User> findByUsername(String username);
    boolean existsByEmail(String email);
    boolean existsByUsername(String username);

    @Query("SELECT u FROM User u WHERE LOWER(u.username) LIKE LOWER(CONCAT(:query, '%')) ORDER BY u.username")
    List<User> searchByUsername(@Param("query") String query);

    @Query("SELECT u FROM User u WHERE LOWER(u.username) LIKE :query OR LOWER(u.email) LIKE :query ORDER BY u.username")
    List<User> searchByUsernameOrEmail(@Param("query") String query);
}
