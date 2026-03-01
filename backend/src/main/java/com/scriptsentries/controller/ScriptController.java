package com.scriptsentries.controller;

import com.scriptsentries.dto.RiskFlagResponse;
import com.scriptsentries.dto.RiskUpdateRequest;
import com.scriptsentries.dto.ScriptResponse;
import com.scriptsentries.model.Project;
import com.scriptsentries.model.RiskFlag;
import com.scriptsentries.model.Script;
import com.scriptsentries.repository.ProjectRepository;
import com.scriptsentries.repository.RiskFlagRepository;
import com.scriptsentries.repository.ScriptRepository;
import com.scriptsentries.service.ExcelExportService;
import com.scriptsentries.service.PdfExtractionService;
import com.scriptsentries.service.ScriptAnalysisService;
import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;
import java.util.UUID;

@RestController
@RequestMapping("/api")
@RequiredArgsConstructor
@Slf4j
@CrossOrigin(origins = {"http://localhost:5173", "http://localhost:3000"})
public class ScriptController {

    private final ScriptRepository scriptRepository;
    private final RiskFlagRepository riskFlagRepository;
    private final PdfExtractionService pdfExtractionService;
    private final ScriptAnalysisService scriptAnalysisService;
    private final ExcelExportService excelExportService;
    private final ProjectRepository projectRepository;

    // ─────────────────────────────────────────────────────────────────────────
    // GET /api/scripts — list all scripts (summaries)
    // ─────────────────────────────────────────────────────────────────────────
    @GetMapping("/scripts")
    public ResponseEntity<List<ScriptResponse>> listScripts() {
        List<ScriptResponse> summaries = scriptRepository.findAllByOrderByUploadedAtDesc()
                .stream()
                .map(ScriptResponse::summary)
                .toList();
        return ResponseEntity.ok(summaries);
    }

    // ─────────────────────────────────────────────────────────────────────────
    // GET /api/scripts/{id} — get script with all risk flags
    // ─────────────────────────────────────────────────────────────────────────
    @GetMapping("/scripts/{id:\\d+}")
    public ResponseEntity<ScriptResponse> getScript(@PathVariable Long id) {
        Script script = scriptRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Script not found: " + id));

        List<RiskFlagResponse> risks = riskFlagRepository.findByScriptOrderBySeverityAsc(script)
                .stream()
                .map(RiskFlagResponse::from)
                .toList();

        return ResponseEntity.ok(ScriptResponse.from(script, risks));
    }

    // ─────────────────────────────────────────────────────────────────────────
    // POST /api/scripts/scan — ZERO RETENTION PDF ANALYSIS (PROJECT-AWARE)
    // ─────────────────────────────────────────────────────────────────────────
    @PostMapping(value = "/scripts/scan", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<ScriptResponse> scanScript(
            @RequestParam("file") MultipartFile file,
            @RequestParam("projectId") Long projectId) throws IOException {

        if (file.isEmpty()) {
            return ResponseEntity.badRequest().build();
        }

        // 1. Verify Project Exists
        Project project = projectRepository.findById(projectId)
                .orElseThrow(() -> new RuntimeException("Project not found: " + projectId));

        String originalFilename = file.getOriginalFilename() != null ? file.getOriginalFilename() : "script.pdf";
        log.info("Received script for analysis: {} under project: {}", originalFilename, project.getName());

        // 2. Save script metadata LINKED TO PROJECT
        Script script = Script.builder()
                .filename(originalFilename)
                .totalPages(0)
                .riskCount(0)
                .project(project)
                .status(Script.ScriptStatus.PROCESSING)
                .build();
        script = scriptRepository.save(script);

        // Write PDF to a TEMP FILE — ZERO RETENTION
        File tempFile = File.createTempFile("ss_" + UUID.randomUUID(), ".pdf");

        try {
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                fos.write(file.getBytes());
            }

            List<String> pages = pdfExtractionService.extractPages(tempFile);
            log.info("Extracted {} pages from '{}'", pages.size(), originalFilename);

            script.setTotalPages(pages.size());
            script = scriptRepository.save(script);

            // Run AI analysis
            List<RiskFlag> flags = scriptAnalysisService.analyzeScript(pages, script);
            riskFlagRepository.saveAll(flags);

            // Update script with final counts and status
            script.setRiskCount(flags.size());
            script.setStatus(Script.ScriptStatus.COMPLETE);
            script = scriptRepository.save(script);

            log.info("Analysis complete for '{}': {} risks found", originalFilename, flags.size());

            List<RiskFlagResponse> risks = flags.stream().map(RiskFlagResponse::from).toList();
            return ResponseEntity.ok(ScriptResponse.from(script, risks));

        } catch (Exception e) {
            script.setStatus(Script.ScriptStatus.FAILED);
            scriptRepository.save(script);
            log.error("Analysis failed for '{}': {}", originalFilename, e.getMessage(), e);
            throw e;
        } finally {
            // ZERO-RETENTION: CRITICAL — delete temp PDF regardless of outcome
            boolean deleted = tempFile.delete();
            if (deleted) {
                log.info("ZERO-RETENTION: Temp PDF '{}' securely deleted", tempFile.getName());
            } else {
                log.error("ZERO-RETENTION VIOLATION: Failed to delete temp file '{}'! Manual cleanup required.", tempFile.getAbsolutePath());
                try {
                    Files.deleteIfExists(tempFile.toPath());
                } catch (IOException ex) {
                    log.error("NIO fallback delete also failed: {}", ex.getMessage());
                }
            }
        }
    }

    // ─────────────────────────────────────────────────────────────────────────
    // DELETE /api/scripts/{id} — Soft Delete Clearance Report
    // ─────────────────────────────────────────────────────────────────────────
    @DeleteMapping("/scripts/{id:\\d+}")
    public ResponseEntity<Void> deleteScript(
            @PathVariable Long id,
            @RequestParam("userId") Long userId) {

        log.info("Request to delete script ID {} by user ID {}", id, userId);

        Script script = scriptRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Script not found"));

        script.setDeletedAt(LocalDateTime.now());
        scriptRepository.save(script);

        return ResponseEntity.noContent().build();
    }

    // ─────────────────────────────────────────────────────────────────────────
    // PATCH /api/risks/{id} — update risk flag (status, comments, redaction)
    // ─────────────────────────────────────────────────────────────────────────
    @PatchMapping("/risks/{id}")
    public ResponseEntity<RiskFlagResponse> updateRisk(
            @PathVariable Long id,
            @RequestBody RiskUpdateRequest request) {

        RiskFlag flag = riskFlagRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Risk flag not found: " + id));

        if (request.getStatus() != null) {
            flag.setStatus(request.getStatus());
        }
        if (request.getComments() != null) {
            flag.setComments(request.getComments());
        }
        if (request.getRestrictions() != null) {
            flag.setRestrictions(request.getRestrictions());
        }
        if (request.getIsRedacted() != null) {
            flag.setRedacted(request.getIsRedacted());
            log.info("Redaction {} on risk flag {} (entity: '{}')",
                    request.getIsRedacted() ? "ENABLED" : "DISABLED",
                    id, flag.getEntityName());
        }

        flag = riskFlagRepository.save(flag);
        return ResponseEntity.ok(RiskFlagResponse.from(flag));
    }

    // ─────────────────────────────────────────────────────────────────────────
    // GET /api/scripts/{id}/export — Secure Excel export with redaction
    // ─────────────────────────────────────────────────────────────────────────
    @GetMapping("/scripts/{id:\\d+}/export")
    public ResponseEntity<byte[]> exportScript(@PathVariable Long id) throws IOException {
        Script script = scriptRepository.findById(id)
                .orElseThrow(() -> new RuntimeException("Script not found: " + id));

        List<RiskFlag> risks = riskFlagRepository.findByScriptOrderBySeverityAsc(script);

        byte[] excelBytes = excelExportService.generateReport(script, risks);

        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmm"));
        String safeFilename = script.getFilename().replaceAll("[^a-zA-Z0-9._-]", "_").replace(".pdf", "");
        String exportFilename = "ScriptSentries_" + safeFilename + "_" + timestamp + ".xlsx";

        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.parseMediaType("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
        headers.setContentDisposition(ContentDisposition.attachment().filename(exportFilename).build());
        headers.setContentLength(excelBytes.length);

        long redactedCount = risks.stream().filter(RiskFlag::isRedacted).count();
        log.info("Excel export for '{}': {} risks, {} redacted rows", script.getFilename(), risks.size(), redactedCount);

        return ResponseEntity.ok().headers(headers).body(excelBytes);
    }
}