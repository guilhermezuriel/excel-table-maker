package com.guilhermezuriel.exceltablemaker.excelGenerator.base;

import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

public abstract class BaseExcel implements BaseExcelService {
    protected byte[] createExcelSheet(String fileName, AbstractList<?> data, Set<String> columns) throws IOException {
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
            applyTilte(sheet, workbook, rowCount++, name);
            applyHeaderRows(workbook, sheet, rowCount++, columns);
            XSSFCellStyle dataCellStyle = ExcelStyle.createDataStyle(workbook);
            applyDataToSheet(data, columns, workbook, sheet, rowCount++, dataCellStyle);

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

    private void applyHeaderRows(XSSFWorkbook workbook, XSSFSheet sheet, int rowCount, Set<String> columns) {
        XSSFCellStyle headerCellStyle = ExcelStyle.createHeaderStyle(workbook);
        XSSFRow headerRow = sheet.createRow(rowCount);
        headerRow.setHeightInPoints((short) 20);
        var iterator = columns.iterator();
        for (int i = 0; i < columns.size(); i++) {
            XSSFCell cell = headerRow.createCell(i);
            cell.setCellValue(iterator.next());
            cell.setCellStyle(headerCellStyle);
        }
    }

    private void applyTilte(XSSFSheet sheet, XSSFWorkbook workbook, int rowCount, String titleName) {
        XSSFCellStyle titleStyle = ExcelStyle.createTitleStyle(workbook);
        XSSFRow titleRow = sheet.createRow(rowCount);
        titleRow.setHeightInPoints((short) 50);
        XSSFCell title = titleRow.createCell(0);
        title.setCellStyle(titleStyle);
        title.setCellValue(titleName);
    }

}
