package com.guilhermezuriel.exceltablemaker.service.excel.impl.excelGenerator.impl;

import com.guilhermezuriel.exceltablemaker.service.excel.impl.excelGenerator.base.BaseExcel;
import com.guilhermezuriel.exceltablemaker.service.excel.dtos.StyleExcelTable;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.AbstractList;
import java.util.LinkedHashMap;
import java.util.Set;
import java.util.UUID;

@Component
public class ExcelGeneratorWeb extends BaseExcel{

    /**
     * Method to create an Excel spreadsheet. The data can be loaded from a Request Object.
     * The columns name are extracted from the object data field;
     *
     * @param data - List containing the data to be filled
     * @param name -  Table name
     * @param style - Personalized style object
     */
    @Override
    public byte[] generateExcelTable(AbstractList<?> data, String name, StyleExcelTable style) throws IOException {
        Set<String> columns = extractColumnsByReference(data);
       return this.createExcelSheet(name, data, columns, style);
    }

    @Override
    public byte[] generateExcelTable(AbstractList<?> data, StyleExcelTable style) throws IOException {
        String name = UUID.randomUUID().toString();
        return generateExcelTable(data, name, style);
    }

    @Override
    public void applyDataToSheet(AbstractList<?> data, Set<String> columns, XSSFWorkbook workbook, XSSFSheet sheet, int rowCount, XSSFCellStyle style) {
        for (Object object : data) {
            XSSFRow row = sheet.createRow(rowCount++);
            LinkedHashMap<String, Object> map = (LinkedHashMap<String, Object>) object;
            var newIterator = columns.iterator();
            for (int i = 0; i < columns.size(); i++) {
                XSSFCell cell = row.createCell(i);
                Object value = map.get(newIterator.next());
                this.setDataCellWithSyle(style, cell, value);
            }
        }
    }

    private Set<String> extractColumnsByReference(AbstractList<?> list) {
        var columsnReference = (LinkedHashMap<String, Object>) list.getFirst();
        return columsnReference.keySet();
    }

}
