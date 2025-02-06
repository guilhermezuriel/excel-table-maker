package com.guilhermezuriel.exceltablemaker.excelGenerator.impl;

import com.guilhermezuriel.exceltablemaker.excelGenerator.annotations.ExcelColumn;
import com.guilhermezuriel.exceltablemaker.excelGenerator.annotations.ExcelTable;
import com.guilhermezuriel.exceltablemaker.excelGenerator.base.BaseExcel;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.lang.reflect.Field;
import java.util.AbstractList;
import java.util.HashSet;
import java.util.Optional;
import java.util.Set;

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
    public byte[] generateExcelTable(AbstractList<?> data) throws IOException {
        validateDataList(data);
        String fileName = extractSheetNameByAnnotations(data);
        Set<String> columns = exctractColumnsByAnnotations(data);
        return this.createExcelSheet(fileName, data, columns);
    }

    @Override
    public byte[] generateExcelTable(AbstractList<?> data, String sheetName) throws IOException {
        validateDataList(data);
        Set<String> columns = exctractColumnsByAnnotations(data);
        return this.createExcelSheet(sheetName, data, columns);
    }

    @Override
    public void applyDataToSheet(AbstractList<?> data, Set<String> columns, XSSFWorkbook workbook, XSSFSheet sheet, int rowCount, XSSFCellStyle cellStyle) {

    }

    private void validateDataList(AbstractList<?> data) {
        if (data.isEmpty()) {
            throw new RuntimeException();
        }
        Object first = data.getFirst();
        Class<?> aClass = first.getClass();
        if (aClass.getAnnotation(ExcelTable.class) == null) {
            throw new RuntimeException();
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

    private Object getFieldByIndex(Object object, Field[] fields, int index) throws IllegalAccessException {
        if (index < 0 || index >= fields.length) {
            throw new IndexOutOfBoundsException("Index out of range");
        }
        fields[index].setAccessible(true);
        return fields[index].get(object);
    }

}
