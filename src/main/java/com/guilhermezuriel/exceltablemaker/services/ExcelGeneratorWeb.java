package com.guilhermezuriel.exceltablemaker.services;

import com.guilhermezuriel.exceltablemaker.annotations.ExcelColumn;
import com.guilhermezuriel.exceltablemaker.annotations.ExcelTable;
import com.guilhermezuriel.exceltablemaker.dtos.RequestExcelTable;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;


@Service
public class ExcelGeneratorWeb {

    public byte[] criarPlanilhaExcel(RequestExcelTable request) throws IOException {
//        this.validarListaDeDados(request.getData());

        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            //TODO: CHANGE
            String nomePlanilha = request.getName() + ".xlsx";
            Set<String> nomeColunas = this.extrairColunasPorAnotacoes(request.getData());

            if (nomeColunas.isEmpty()) {
                return null;
            }

            int rowCount = 0;

            var data = request.getData();

            XSSFCellStyle titleStyle = this.criarEstiloDoTitulo(workbook);
            XSSFCellStyle headerCellStyle = this.criarEstiloHeaderDoExcel(workbook);
            XSSFCellStyle dataCellStyle = this.criarEstiloDosDados(workbook);

            XSSFSheet sheet = workbook.createSheet(nomePlanilha);

            XSSFRow titleRow = sheet.createRow(rowCount++);
            titleRow.setHeightInPoints((short) 50);
            XSSFCell titulo = titleRow.createCell(0);
            titulo.setCellStyle(titleStyle);
            titulo.setCellValue(request.getName());

            XSSFRow headerRow = sheet.createRow(rowCount++);
            headerRow.setHeightInPoints((short) 20);
           var iterator = nomeColunas.iterator();
            for (int i = 0; i < nomeColunas.size(); i++) {
                XSSFCell cell = headerRow.createCell(i);
                cell.setCellValue(iterator.next());
                cell.setCellStyle(headerCellStyle);
            }

            for (Object object : data) {
                XSSFRow row = sheet.createRow(rowCount++);
                LinkedHashMap<String, Object> map = (LinkedHashMap<String, Object>) object;
                var newIterator = nomeColunas.iterator();
                for (int i = 0; i < nomeColunas.size(); i++) {
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
            for (int i = 0; i < nomeColunas.size(); i++) {
                sheet.autoSizeColumn(i);
            }
            try (ByteArrayOutputStream fileOutputStream = new ByteArrayOutputStream()) {
                workbook.write(fileOutputStream);
                return fileOutputStream.toByteArray();
            }
        }
    }


    private void validarListaDeDados(List<?> listaDados) {
        if (listaDados.isEmpty()) {
            throw new RuntimeException();
        }
        Object first = listaDados.getFirst();
        Class<?> aClass = first.getClass();
        if (aClass.getAnnotation(ExcelTable.class) == null) {
            throw new RuntimeException();
        }
    }


    private String extrairNomePlanilhaPorAnotacao(Object object) {
        String nomePlanilha = object.getClass().getAnnotation(ExcelTable.class).name();
        if (nomePlanilha == null || nomePlanilha.isEmpty()) return object.getClass().getSimpleName();
        return nomePlanilha;
    }

    private Set<String> extrairColunasPorAnotacoes(List<?> list) {
        LinkedHashMap<String, String> object = (LinkedHashMap<String, String>) list.getFirst();
        Set<String> fields = object.keySet();
        return fields;
    }

    private XSSFCellStyle criarEstiloDoTitulo(XSSFWorkbook workbook) {
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


    private XSSFCellStyle criarEstiloHeaderDoExcel(XSSFWorkbook workbook) {
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

    private XSSFCellStyle criarEstiloDosDados(XSSFWorkbook workbook) {
        XSSFCellStyle dataCellStyle = workbook.createCellStyle();
        dataCellStyle.setAlignment(HorizontalAlignment.CENTER);
        dataCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dataCellStyle.setWrapText(true);
        dataCellStyle.setBorderRight(BorderStyle.THIN);
        dataCellStyle.setBorderTop(BorderStyle.THIN);
        dataCellStyle.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
        return dataCellStyle;
    }

    private Object getFieldByIndex(Object object, Field[] fields, int index) throws IllegalAccessException {
        if (index < 0 || index >= fields.length) {
            throw new IndexOutOfBoundsException("Index out of range");
        }
        fields[index].setAccessible(true);
        return fields[index].get(object);
    }
}
