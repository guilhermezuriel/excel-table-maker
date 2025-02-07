package com.guilhermezuriel.exceltablemaker.excelGenerator.impl;

import com.guilhermezuriel.exceltablemaker.excelGenerator.annotations.ExcelColumn;
import com.guilhermezuriel.exceltablemaker.excelGenerator.annotations.ExcelTable;
import com.guilhermezuriel.exceltablemaker.excelGenerator.base.BaseExcel;
import com.guilhermezuriel.exceltablemaker.service.dtos.StyleExcelTable;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.*;

@Component
public class ExcelGeneratorLocal extends BaseExcel{
    /**
     * Method to create an Excel spreadsheet. The data can be loaded from any DTO, as long as it is not a wrapper class.
     * The sheet name is retrieved through the @ExcelTable annotation, present in the data class.
     * The column names are directly related to the @ExcelColumn annotation.
     *
     * @param data - List containing the data to be filled
     */
    @Override
    public byte[] generateExcelTable(AbstractList<?> data, StyleExcelTable style) throws IOException {
        validateDataList(data);
        String fileName = extractSheetNameByAnnotations(data);
        Set<String> columns = exctractColumnsByAnnotations(data);
        return this.createExcelSheet(fileName, data, columns, style);
    }

    @Override
    public byte[] generateExcelTable(AbstractList<?> data, String sheetName, StyleExcelTable style) throws IOException {
        validateDataList(data);
        Set<String> columns = exctractColumnsByAnnotations(data);
        return this.createExcelSheet(sheetName, data, columns, style);
    }

    @Override
    public void applyDataToSheet(AbstractList<?> data, Set<String> columns, XSSFWorkbook workbook, XSSFSheet sheet, int rowCount, XSSFCellStyle style)  {
        Field[] fields = Arrays.stream(data.getFirst().getClass().getDeclaredFields()).filter(field -> field.isAnnotationPresent(ExcelColumn.class)).toArray(Field[]::new);
        for (Object object : data) {
            XSSFRow row = sheet.createRow(rowCount++);
            for (int i = 0; i < fields.length; i++) {
                XSSFCell cell = row.createCell(i);
                Object value = this.getFieldByIndex(object, fields, i);
                this.setDataCellWithSyle(style, cell, value);
            }
        }
    }

    private void validateDataList(AbstractList<?> data) {
        if (data.isEmpty()) {
            throw new RuntimeException();
        }
        Object first = data.getFirst();
        boolean instance = first instanceof Class;
        if(!instance) {
            throw new IllegalArgumentException("Data is not a class");
        }
        Class<?> aClass = first.getClass();
        if (aClass.getAnnotation(ExcelTable.class) == null) {
            throw new RuntimeException("Data is not annotated with @ExcelTable");
        }
    }


    private String extractSheetNameByAnnotations(Object object) {
        String nomePlanilha = object.getClass().getAnnotation(ExcelTable.class).name();
        if (nomePlanilha == null || nomePlanilha.isEmpty()) return object.getClass().getSimpleName();
        return nomePlanilha;
    }

    private Set<String> exctractColumnsByAnnotations(Object object) {
        Field[] fields = object.getClass().getDeclaredFields();
        Set<String> columns = new HashSet<>();
        for (Field field : fields) {
            Optional<ExcelColumn> annotation = Optional.ofNullable(field.getAnnotation(ExcelColumn.class));
            if (annotation.isPresent()) {
                String definedColumnName = annotation.get().name();
                if (definedColumnName.isEmpty()) {
                    columns.add(field.getName());
                    continue;
                }
                columns.add(definedColumnName);
            }
        }
        return columns;
    }

    private Object getFieldByIndex(Object object, Field[] fields, int index) {
        try {
            if (index < 0 || index >= fields.length) {
                throw new IndexOutOfBoundsException("Index out of range");
            }
            fields[index].setAccessible(true);
            return fields[index].get(object);
        }catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }

}
