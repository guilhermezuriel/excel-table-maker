package com.guilhermezuriel.exceltablemaker.service.excel.impl.excelGenerator.impl;

import com.guilhermezuriel.exceltablemaker.service.excel.dtos.StyleExcelTable;
import com.guilhermezuriel.exceltablemaker.service.excel.impl.excelGenerator.base.BaseExcel;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.AbstractList;
import java.util.Set;

@Component
public class ExcelGeneratorCsv extends BaseExcel {


    @Override
    public void applyDataToSheet(AbstractList<?> data, Set<String> columns, XSSFWorkbook workbook, XSSFSheet sheet, int rowCount, XSSFCellStyle cellStyle) {
        //TODO: Ao contrario dos outros, esse metodo sera aplicado recursivamente, ao inves dele proprio conter um laço
    }

    @Override
    public byte[] generateExcelTable(AbstractList<?> data, StyleExcelTable style) throws IOException {
        //TODO: extrair colunas a partir da primeira lista / primeira linha do arquivo
        //TODO: Listas seram matrizes, extrair cada matriz e ir aplicando os dados ao ler cada lista dentro de data
        return new byte[0];
    }

    @Override
    public byte[] generateExcelTable(AbstractList<?> data, String sheetName, StyleExcelTable style) throws IOException {
        return new byte[0];
    }
}
