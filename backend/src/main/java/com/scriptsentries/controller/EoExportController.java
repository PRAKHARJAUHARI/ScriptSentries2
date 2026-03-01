package com.scriptsentries.controller;

import com.scriptsentries.service.EoExportService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

@RestController
@RequestMapping("/api/scripts")
@RequiredArgsConstructor
@Slf4j
@CrossOrigin(origins = {"http://localhost:5173", "http://localhost:3000"})
public class EoExportController {

    private final EoExportService eoExportService;

    /**
     * GET /api/scripts/{id}/export/eo?userId=7
     *
     * Generates and downloads the E&O insurance Excel workbook for the given script.
     * Contains 3 tabs:
     *   1. E&O Cover Sheet  — production info + attorney certification + statistics
     *   2. Full Risk Log    — all risks with US legal exposure categories
     *   3. Cleared Items    — insurer confirmation sheet (cleared items only)
     *
     * Requires: requesting user must be ATTORNEY or ANALYST on the project.
     */
    @GetMapping("/{id}/export/eo")
    public ResponseEntity<byte[]> exportEo(
            @PathVariable Long id,
            @RequestParam Long userId) throws IOException {

        byte[] excelBytes = eoExportService.generateEoReport(id, userId);

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmm"));
        String filename  = "EO_Submission_" + id + "_" + timestamp + ".xlsx";

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDisposition(ContentDisposition.attachment().filename(filename).build());
        headers.setContentLength(excelBytes.length);

        return ResponseEntity.ok().headers(headers).body(excelBytes);
    }
}
