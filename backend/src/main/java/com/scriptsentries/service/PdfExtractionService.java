package com.scriptsentries.service;

import lombok.extern.slf4j.Slf4j;
import org.apache.pdfbox.Loader;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Extracts text content from a PDF file page-by-page using Apache PDFBox 3.0.
 * Designed for use in conjunction with Zero-Retention policy — callers must
 * ensure the temp file is deleted in a try-finally block.
 */
@Service
@Slf4j
public class PdfExtractionService {

    /**
     * Extracts all text pages from a PDF file.
     *
     * @param pdfFile the temp file (caller is responsible for deletion)
     * @return ordered list of page text strings (index 0 = page 1)
     */
    public List<String> extractPages(File pdfFile) throws IOException {
        List<String> pages = new ArrayList<>();

        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            int totalPages = document.getNumberOfPages();
            log.info("PDF loaded: {} pages", totalPages);

            PDFTextStripper stripper = new PDFTextStripper();

            for (int i = 1; i <= totalPages; i++) {
                stripper.setStartPage(i);
                stripper.setEndPage(i);
                String pageText = stripper.getText(document).trim();
                pages.add(pageText);
            }
        }

        return pages;
    }

    /**
     * Returns total page count without extracting content.
     */
    public int getPageCount(File pdfFile) throws IOException {
        try (PDDocument document = Loader.loadPDF(pdfFile)) {
            return document.getNumberOfPages();
        }
    }
}
