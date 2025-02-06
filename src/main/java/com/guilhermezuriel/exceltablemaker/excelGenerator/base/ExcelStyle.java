package com.guilhermezuriel.exceltablemaker.excelGenerator.base;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelStyle {

    public static XSSFCellStyle createTitleStyle(XSSFWorkbook workbook) {
        Font titleFont = workbook.createFont();
        titleFont.setFontHeightInPoints((short) 36);
        titleFont.setColor(IndexedColors.DARK_BLUE.getIndex());

        XSSFCellStyle titleCellStyle = workbook.createCellStyle();
        return getXssfCellStyle(titleFont, titleCellStyle);
    }

    public static XSSFCellStyle createHeaderStyle(XSSFWorkbook workbook) {
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 13);
        headerFont.setColor(IndexedColors.WHITE.getIndex());
        headerFont.setItalic(false);
        return getXssfCellStyle(workbook, headerFont);
    }

    public static XSSFCellStyle createDataStyle(XSSFWorkbook workbook) {
        return getXssfCellStyle(workbook);
    }

    private static XSSFCellStyle getXssfCellStyle(Font titleFont, XSSFCellStyle titleCellStyle) {
        titleCellStyle.setFont(titleFont);
        titleCellStyle.setAlignment(HorizontalAlignment.LEFT);
        titleCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleCellStyle.setBorderBottom(BorderStyle.NONE);
        titleCellStyle.setBorderLeft(BorderStyle.NONE);
        titleCellStyle.setBorderRight(BorderStyle.NONE);
        titleCellStyle.setBorderTop(BorderStyle.NONE);

        return titleCellStyle;
    }

    private static XSSFCellStyle getXssfCellStyle(XSSFWorkbook workbook, Font headerFont) {
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

    private static XSSFCellStyle getXssfCellStyle(XSSFWorkbook workbook) {
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
