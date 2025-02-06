package com.guilhermezuriel.exceltablemaker.excelGenerator.impl;

import com.guilhermezuriel.exceltablemaker.excelGenerator.base.BaseExcel;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.AbstractList;
import java.util.LinkedHashMap;
import java.util.Set;
import java.util.UUID;

@Component
public class ExcelGeneratorWeb extends BaseExcel{

    @Override
    public byte[] generateExcelTable(AbstractList<?> data, String name) throws IOException {
        Set<String> columns = extractColumnsByReference((AbstractList<?>) data);
       return this.createExcelSheet(name, data, columns);
    }

    @Override
    public byte[] generateExcelTable(AbstractList<?> data) throws IOException {
        String name = UUID.randomUUID().toString();
        return generateExcelTable(data, name);
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
        }
    }

    private Set<String> extractColumnsByReference(AbstractList<?> list) {
        var columsnReference = (LinkedHashMap<String, Object>) list.getFirst();
        return columsnReference.keySet();
    }

}
