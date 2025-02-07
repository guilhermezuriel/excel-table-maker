package com.guilhermezuriel.exceltablemaker.excelGenerator.base;

import com.guilhermezuriel.exceltablemaker.excelGenerator.style.CellStyle;
import com.guilhermezuriel.exceltablemaker.service.dtos.StyleExcelTable;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.AbstractList;
import java.util.Objects;
import java.util.Set;
import java.util.UUID;

public abstract class BaseExcel implements BaseExcelService {
    protected byte[] createExcelSheet(String fileName, AbstractList<?> data, Set<String> columns, StyleExcelTable style) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            String name = fileName != null ? fileName : "Sheet - "+ UUID.randomUUID();
            if(Objects.isNull(data) || data.isEmpty()) {
                throw new RuntimeException("Data is empty");
            }
            if (columns.isEmpty()) {
                return null;
            }
            int rowCount = 0;
            XSSFSheet sheet = workbook.createSheet(name);
            applyTilte(sheet, workbook, rowCount++, name, style.getTitle());
            applyHeaderRows(workbook, sheet, rowCount++, columns, style.getHeader());
            XSSFCellStyle dataCellStyle = ExcelStyle.createStyle(workbook, style.getData());
            applyDataToSheet(data, columns, workbook, sheet, rowCount++, dataCellStyle);

            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, columns.size() - 1));
            for (int i = 0; i <= columns.size() - 1; i++) {
                sheet.setColumnWidth(i, 4000);
            }

            for (int i = 0; i < columns.size(); i++) {
                sheet.autoSizeColumn(i);
            }

            try (ByteArrayOutputStream fileOutputStream = new ByteArrayOutputStream()) {
                workbook.write(fileOutputStream);
                return fileOutputStream.toByteArray();
            }
        }
    }

    protected void setDataCellWithSyle(XSSFCellStyle style, XSSFCell cell, Object value) {
        cell.setCellStyle(style);
        switch (value) {
            case String s -> cell.setCellValue(s);
            case Integer j -> cell.setCellValue(j);
            case Boolean b -> cell.setCellValue(b);
            case Enum<?> e -> cell.setCellValue(e.name());
            case BigDecimal bd -> cell.setCellValue(bd.doubleValue());
            case Double d -> cell.setCellValue(d);
            case LocalDate ld -> cell.setCellValue(ld.toString());
            case LocalDateTime ldt -> cell.setCellValue(ldt.toString());
            default -> throw new IllegalArgumentException("Unsupported type: Cell Value must be a primitive type" + value.getClass());
        }
    }

    private void applyHeaderRows(XSSFWorkbook workbook, XSSFSheet sheet, int rowCount, Set<String> columns, CellStyle style) {
        XSSFCellStyle headerCellStyle = ExcelStyle.createStyle(workbook, style);
        XSSFRow headerRow = sheet.createRow(rowCount);
        headerRow.setHeightInPoints((short) 20);
        var iterator = columns.iterator();
        for (int i = 0; i < columns.size(); i++) {
            XSSFCell cell = headerRow.createCell(i);
            cell.setCellValue(iterator.next());
            cell.setCellStyle(headerCellStyle);
        }
    }

    private void applyTilte(XSSFSheet sheet, XSSFWorkbook workbook, int rowCount, String titleName, CellStyle style) {
        XSSFCellStyle titleStyle = ExcelStyle.createStyle(workbook, style);
        XSSFRow titleRow = sheet.createRow(rowCount);
        titleRow.setHeightInPoints((short) 20);
        XSSFCell title = titleRow.createCell(rowCount, CellType.STRING);
        title.setCellStyle(titleStyle);
        title.setCellValue(titleName);
    }

}
