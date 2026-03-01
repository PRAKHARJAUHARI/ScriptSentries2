package com.scriptsentries.repository;

import com.scriptsentries.model.Comment;
import com.scriptsentries.model.RiskFlag;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface CommentRepository extends JpaRepository<Comment, Long> {
    List<Comment> findByRiskFlagOrderByCreatedAtAsc(RiskFlag riskFlag);
}
