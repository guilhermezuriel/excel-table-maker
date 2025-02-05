package com.guilhermezuriel.exceltablemaker.services;

import org.apache.poi.xssf.usermodel.*;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Objects;
import java.util.Set;
import java.util.UUID;

public abstract class ExcelGeneratorInterface{
    protected byte[] createExcelSheet(String fileName, List<?> data,  Set<String> columns) throws IOException {
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

    public void applyDataToSheet(List<?> data, Set<String> columns, XSSFWorkbook workbook, XSSFSheet sheet, int rowCount, XSSFCellStyle cellStyle) {
        throw new UnsupportedOperationException("This method needs to be implemented by his siblings.");
    }
}
