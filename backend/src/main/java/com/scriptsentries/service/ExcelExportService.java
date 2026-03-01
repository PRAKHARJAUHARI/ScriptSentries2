package com.scriptsentries.service;

import com.scriptsentries.model.RiskFlag;
import com.scriptsentries.model.Script;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.time.format.DateTimeFormatter;
import java.util.List;

/**
 * Generates a sanitized Excel clearance report.
 * Rows with isRedacted=true have sensitive columns replaced with [REDACTED].
 */
@Service
@Slf4j
public class ExcelExportService {

    private static final String REDACTED = "[REDACTED]";
    private static final DateTimeFormatter DATE_FMT = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm");

    public byte[] generateReport(Script script, List<RiskFlag> risks) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Clearance Report");

            // Styles
            CellStyle headerStyle = createHeaderStyle(workbook);
            CellStyle titleStyle = createTitleStyle(workbook);
            CellStyle highStyle = createSeverityStyle(workbook, new XSSFColor(new byte[]{(byte)220, (byte)53, (byte)69}, null));
            CellStyle medStyle = createSeverityStyle(workbook, new XSSFColor(new byte[]{(byte)255, (byte)193, (byte)7}, null));
            CellStyle lowStyle = createSeverityStyle(workbook, new XSSFColor(new byte[]{(byte)25, (byte)135, (byte)84}, null));
            CellStyle redactedStyle = createRedactedStyle(workbook);
            CellStyle dataStyle = createDataStyle(workbook);

            int rowIdx = 0;

            // Title block
            Row titleRow = sheet.createRow(rowIdx++);
            Cell titleCell = titleRow.createCell(0);
            titleCell.setCellValue("SCRIPTSENTRIES — LEGAL CLEARANCE REPORT");
            titleCell.setCellStyle(titleStyle);
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 11));

            Row metaRow = sheet.createRow(rowIdx++);
            metaRow.createCell(0).setCellValue("Script: " + script.getFilename());
            metaRow.createCell(4).setCellValue("Pages: " + script.getTotalPages());
            metaRow.createCell(6).setCellValue("Risks: " + risks.size());
            metaRow.createCell(8).setCellValue("Generated: " +
                    (script.getUploadedAt() != null ? script.getUploadedAt().format(DATE_FMT) : "N/A"));

            rowIdx++; // spacer

            // Column headers
            String[] headers = {
                "Page", "Severity", "Category", "Sub-Category",
                "Entity Name", "Snippet", "Reason", "Suggestion",
                "Status", "Comments", "Restrictions", "Redacted"
            };
            Row headerRow = sheet.createRow(rowIdx++);
            for (int col = 0; col < headers.length; col++) {
                Cell cell = headerRow.createCell(col);
                cell.setCellValue(headers[col]);
                cell.setCellStyle(headerStyle);
            }

            // Data rows
            for (RiskFlag risk : risks) {
                Row row = sheet.createRow(rowIdx++);

                CellStyle severityStyle = switch (risk.getSeverity()) {
                    case HIGH -> highStyle;
                    case MEDIUM -> medStyle;
                    case LOW -> lowStyle;
                };

                setCell(row, 0, String.valueOf(risk.getPageNumber()), dataStyle);
                setCell(row, 1, risk.getSeverity().name(), severityStyle);
                setCell(row, 2, risk.getCategory().name(), dataStyle);
                setCell(row, 3, risk.getSubCategory().name(), dataStyle);

                if (risk.isRedacted()) {
                    // SECURITY: Replace sensitive fields with [REDACTED]
                    setCell(row, 4, REDACTED, redactedStyle);
                    setCell(row, 5, REDACTED, redactedStyle);
                    setCell(row, 6, risk.getReason(), dataStyle);       // reason stays
                    setCell(row, 7, risk.getSuggestion(), dataStyle);   // suggestion stays
                    setCell(row, 8, risk.getStatus().name(), dataStyle);
                    setCell(row, 9, REDACTED, redactedStyle);
                    setCell(row, 10, REDACTED, redactedStyle);
                    setCell(row, 11, "YES", redactedStyle);
                } else {
                    setCell(row, 4, risk.getEntityName(), dataStyle);
                    setCell(row, 5, risk.getSnippet(), dataStyle);
                    setCell(row, 6, risk.getReason(), dataStyle);
                    setCell(row, 7, risk.getSuggestion(), dataStyle);
                    setCell(row, 8, risk.getStatus().name(), dataStyle);
                    setCell(row, 9, risk.getComments(), dataStyle);
                    setCell(row, 10, risk.getRestrictions(), dataStyle);
                    setCell(row, 11, "NO", dataStyle);
                }
            }

            // Auto-size columns
            int[] colWidths = {8, 12, 22, 28, 25, 40, 45, 45, 25, 35, 35, 12};
            for (int i = 0; i < colWidths.length; i++) {
                sheet.setColumnWidth(i, colWidths[i] * 256);
            }

            ByteArrayOutputStream out = new ByteArrayOutputStream();
            workbook.write(out);
            log.info("Excel report generated: {} rows for script '{}'", risks.size(), script.getFilename());
            return out.toByteArray();
        }
    }

    private void setCell(Row row, int col, String value, CellStyle style) {
        Cell cell = row.createCell(col);
        cell.setCellValue(value != null ? value : "");
        cell.setCellStyle(style);
    }

    private CellStyle createHeaderStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte)15, (byte)23, (byte)42}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont font = wb.createFont();
        font.setBold(true);
        font.setColor(new XSSFColor(new byte[]{(byte)255, (byte)255, (byte)255}, null));
        font.setFontHeightInPoints((short)10);
        style.setFont(font);
        return style;
    }

    private CellStyle createTitleStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte)6, (byte)95, (byte)70}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont font = wb.createFont();
        font.setBold(true);
        font.setColor(new XSSFColor(new byte[]{(byte)255, (byte)255, (byte)255}, null));
        font.setFontHeightInPoints((short)14);
        style.setFont(font);
        return style;
    }

    private CellStyle createSeverityStyle(XSSFWorkbook wb, XSSFColor color) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(color);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont font = wb.createFont();
        font.setBold(true);
        font.setColor(new XSSFColor(new byte[]{(byte)255, (byte)255, (byte)255}, null));
        style.setFont(font);
        return style;
    }

    private CellStyle createRedactedStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(new XSSFColor(new byte[]{(byte)30, (byte)30, (byte)30}, null));
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        XSSFFont font = wb.createFont();
        font.setBold(true);
        font.setColor(new XSSFColor(new byte[]{(byte)255, (byte)80, (byte)80}, null));
        style.setFont(font);
        return style;
    }

    private CellStyle createDataStyle(XSSFWorkbook wb) {
        XSSFCellStyle style = wb.createCellStyle();
        style.setWrapText(true);
        style.setVerticalAlignment(VerticalAlignment.TOP);
        style.setBorderBottom(BorderStyle.HAIR);
        style.setBorderRight(BorderStyle.HAIR);
        return style;
    }
}
