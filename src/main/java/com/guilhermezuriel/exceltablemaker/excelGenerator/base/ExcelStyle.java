package com.guilhermezuriel.exceltablemaker.excelGenerator.base;

import com.guilhermezuriel.exceltablemaker.excelGenerator.style.CellStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Objects;

public class ExcelStyle {

    public static XSSFCellStyle createStyle(XSSFWorkbook workbook, CellStyle cellStyle) {
        Font font = workbook.createFont();
        font.setFontHeightInPoints(cellStyle.fontHeight());
        font.setColor(cellStyle.fontColor().getIndex());
        return personalizeCellStyle(workbook, font, cellStyle);
    }


    private static XSSFCellStyle personalizeCellStyle(XSSFWorkbook workbook, Font font, CellStyle style) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(font);
        cellStyle.setAlignment(style.horizontalAlignment());
        cellStyle.setVerticalAlignment(style.verticalAlignment());
        cellStyle.setWrapText(style.wrapText());

        if(!Objects.isNull(style.borderProps().borderStyle())) {
            cellStyle.setBorderLeft(style.borderProps().borderStyle());
            cellStyle.setBorderRight(style.borderProps().borderStyle());
            cellStyle.setBorderTop(style.borderProps().borderStyle());
            cellStyle.setBorderBottom(style.borderProps().borderStyle());
        }

        if(!Objects.isNull(style.borderProps().borderColor())){
            cellStyle.setLeftBorderColor(style.borderProps().borderColor().getIndex());
            cellStyle.setRightBorderColor(style.borderProps().borderColor().getIndex());
            cellStyle.setTopBorderColor(style.borderProps().borderColor().getIndex());
            cellStyle.setBottomBorderColor(style.borderProps().borderColor().getIndex());
        }

        cellStyle.setFillBackgroundColor(style.backgroundColor().getIndex());
        cellStyle.setFillPattern(style.fillPattern());
        return cellStyle;
    }

}
