package com.scriptsentries.service;

import com.scriptsentries.dto.GlobalSearchDto;
import com.scriptsentries.model.RiskFlag;
import com.scriptsentries.model.User;
import com.scriptsentries.repository.RiskFlagRepository;
import com.scriptsentries.repository.UserRepository;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;

import java.util.List;

@Service
@RequiredArgsConstructor
@Slf4j
public class GlobalSearchService {

    private final RiskFlagRepository riskFlagRepository;
    private final UserRepository     userRepository;

    private static final int MIN_QUERY_LENGTH = 2;
    private static final int MAX_RESULTS      = 200;

    /**
     * Search for an entity name across all projects the user belongs to.
     *
     * @param query  The search term (partial match, case-insensitive)
     * @param userId The requesting user — only their projects are searched
     */
    public GlobalSearchDto.SearchResponse search(String query, Long userId) {
        if (query == null || query.trim().length() < MIN_QUERY_LENGTH) {
            return GlobalSearchDto.SearchResponse.builder()
                    .query(query).totalHits(0).hits(List.of()).build();
        }

        String cleanQuery = query.trim();
        User user = userRepository.findById(userId)
                .orElseThrow(() -> new RuntimeException("User not found: " + userId));

        List<RiskFlag> flags = riskFlagRepository
                .searchByEntityNameAcrossProjects(user, cleanQuery);

        // Cap results to prevent browser overload on very common terms
        List<GlobalSearchDto.SearchHit> hits = flags.stream()
                .limit(MAX_RESULTS)
                .map(this::toHit)
                .toList();

        log.info("Global search '{}' by userId={}: {} hits", cleanQuery, userId, hits.size());

        return GlobalSearchDto.SearchResponse.builder()
                .query(cleanQuery)
                .totalHits(hits.size())
                .hits(hits)
                .build();
    }

    private GlobalSearchDto.SearchHit toHit(RiskFlag r) {
        var script  = r.getScript();
        var project = script != null ? script.getProject() : null;

        return GlobalSearchDto.SearchHit.builder()
                .riskFlagId(r.getId())
                .entityName(r.getEntityName())
                .category(r.getCategory().name())
                .subCategory(r.getSubCategory().name())
                .severity(r.getSeverity().name())
                .status(r.getStatus().name())
                .snippet(r.getSnippet())
                .pageNumber(r.getPageNumber())
                .scriptId(script != null ? script.getId() : null)
                .scriptFilename(script != null ? script.getFilename() : null)
                .versionName(script != null ? script.getVersionName() : null)
                .projectId(project != null ? project.getId() : null)
                .projectName(project != null ? project.getName() : null)
                .studioName(project != null ? project.getStudioName() : null)
                .build();
    }
}