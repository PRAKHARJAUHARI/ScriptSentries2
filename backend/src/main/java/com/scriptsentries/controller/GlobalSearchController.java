package com.scriptsentries.controller;

import com.scriptsentries.dto.GlobalSearchDto;
import com.scriptsentries.service.GlobalSearchService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

@RestController
@RequestMapping("/api/search")
@RequiredArgsConstructor
@CrossOrigin(origins = {"http://localhost:5173", "http://localhost:3000"})
public class GlobalSearchController {

    private final GlobalSearchService searchService;

    /**
     * GET /api/search?q=John+Wick&userId=7
     *
     * Searches entityName across all risk flags in all projects the user belongs to.
     * Returns grouped results with project + script context for each hit.
     *
     * Query must be at least 2 characters.
     * Results are capped at 200 to prevent overload on very broad terms.
     */
    @GetMapping
    public ResponseEntity<GlobalSearchDto.SearchResponse> search(
            @RequestParam String q,
            @RequestParam Long userId) {
        return ResponseEntity.ok(searchService.search(q, userId));
    }
}
