package com.guilhermezuriel.exceltablemaker.services;

import com.guilhermezuriel.exceltablemaker.dtos.RequestExcelTable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.LinkedHashMap;
import java.util.Objects;
import java.util.Set;
import java.util.UUID;


@Service
public class ExcelGeneratorWeb {

    public byte[] createExcelSheet(RequestExcelTable request) throws IOException {
//        this.validarListaDeDados(request.data());
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            //TODO: CHANGE

            String name;
            if(Objects.isNull(request.name())){
                name = UUID.randomUUID().toString();
            }
            name = request.name();
            Set<String> columns = this.extractColumnsByReference((LinkedHashMap<String, Object>) request.data().getFirst());

            if (columns.isEmpty()) {
                return null;
            }

            int rowCount = 0;

            var data = request.data();

            XSSFSheet sheet = workbook.createSheet(name);

           this.createTitle(sheet, workbook, rowCount, name);


            // Headers Rows
            XSSFCellStyle headerCellStyle = this.createHeaderStyle(workbook);
            XSSFRow headerRow = sheet.createRow(rowCount++);
            headerRow.setHeightInPoints((short) 20);
           var iterator = columns.iterator();
            for (int i = 0; i < columns.size(); i++) {
                XSSFCell cell = headerRow.createCell(i);
                cell.setCellValue(iterator.next());
                cell.setCellStyle(headerCellStyle);
            }

            XSSFCellStyle dataCellStyle = this.createDataStyle(workbook);
            for (Object object : data) {
                XSSFRow row = sheet.createRow(rowCount++);
                LinkedHashMap<String, Object> map = (LinkedHashMap<String, Object>) object;
                var newIterator = columns.iterator();
                for (int i = 0; i < columns.size(); i++) {
                    XSSFCell cell = row.createCell(i);
                    Object value = map.get(newIterator.next());
                    cell.setCellStyle(dataCellStyle);
                    switch (value) {
                        case String s -> cell.setCellValue(s);
                        case Integer j -> cell.setCellValue(j);
                        case Boolean b -> cell.setCellValue(b);
                        case Enum<?> e -> cell.setCellValue(e.name());
                        case BigDecimal bd -> cell.setCellValue(bd.doubleValue());
                        case Double d -> cell.setCellValue(d);
                        default -> throw new IllegalArgumentException("Unsupported type: Cell Value must be a primitive type" + value.getClass());
                    }
                }
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

    private Set<String> extractColumnsByReference(LinkedHashMap<String, Object> data) {
        return data.keySet();
    }

    private XSSFCellStyle createTitleStyle(XSSFWorkbook workbook) {
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 36);
        titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());

        XSSFCellStyle titleCellStyle = workbook.createCellStyle();
        titleCellStyle.setFont(titleFont);
        titleCellStyle.setAlignment(HorizontalAlignment.LEFT);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setBorderBottom(BorderStyle.NONE);
        titleCellStyle.setBorderLeft(BorderStyle.NONE);
        titleCellStyle.setBorderRight(BorderStyle.NONE);
        titleCellStyle.setBorderTop(BorderStyle.NONE);

        return titleCellStyle;
    }

    private void createTitle(XSSFSheet sheet, XSSFWorkbook workbook, int rowCount, String titleName) {
        XSSFCellStyle titleStyle = this.createTitleStyle(workbook);
        XSSFRow titleRow = sheet.createRow(rowCount);
        titleRow.setHeightInPoints((short) 50);
        XSSFCell title = titleRow.createCell(0);
        title.setCellStyle(titleStyle);
        title.setCellValue(titleName);
    }


    private XSSFCellStyle createHeaderStyle(XSSFWorkbook workbook) {
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 13);
        headerFont.setColor(IndexedColors.WHITE.getIndex());

        XSSFCellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setFont(headerFont);
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerCellStyle.setWrapText(false);
        headerCellStyle.setBorderBottom(BorderStyle.THIN);
        headerCellStyle.setFillBackgroundColor(IndexedColors.DARK_BLUE.getIndex());
        headerCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return headerCellStyle;
    }

    private XSSFCellStyle createDataStyle(XSSFWorkbook workbook) {
        XSSFCellStyle dataCellStyle = workbook.createCellStyle();
        dataCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dataCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataCellStyle.setWrapText(true);
        dataCellStyle.setBorderRight(BorderStyle.THIN);
        dataCellStyle.setBorderTop(BorderStyle.THIN);
        dataCellStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        return dataCellStyle;
    }

}
