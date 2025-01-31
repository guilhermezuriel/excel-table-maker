package com.guilhermezuriel.exceltablemaker.services;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Optional;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.http.HttpStatus;
import org.springframework.stereotype.Service;

//import br.com.conivales.iconsorcio.application.service.excel.annotations.ExcelColumn;
//import br.com.conivales.iconsorcio.application.service.excel.annotations.ExcelTable;

@Service
public class ExcelGenerator{

    /**
     * Método para criar uma planilha Excel, os dados podem ser carregados a partir de qualquer DTO, contanto que seja não seja uma classe wrapper. O nome da planilha é
     * recuperado através da anotação @ExcelTable, presente na classe dos dados. Os nomes das colunas tem relação direta com @ExcelColumn.
     *
     * @param listaDados - Lista contendo os dados a serem preenchidos
     */
    public byte[] criarPlanilhaExcel(List<?> listaDados, String nomeTitulo) throws IOException {
        this.validarListaDeDados(listaDados);
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            String nomePlanilha = this.extrairNomePlanilhaPorAnotacao(listaDados.get(0));
            List<String> nomeColunas = this.extrairColunasPorAnotacoes(listaDados.get(0));

            if (nomeColunas.isEmpty()) {
                return null;
            }

            int rowCount = 0;

            XSSFCellStyle titleStyle = this.criarEstiloDoTitulo(workbook);
            XSSFCellStyle headerCellStyle = this.criarEstiloHeaderDoExcel(workbook);
            XSSFCellStyle dataCellStyle = this.criarEstiloDosDados(workbook);

            XSSFSheet sheet = workbook.createSheet(nomePlanilha);

            XSSFRow titleRow = sheet.createRow(rowCount++);
            titleRow.setHeightInPoints((short) 50);
            XSSFCell titulo = titleRow.createCell(0);
            titulo.setCellStyle(titleStyle);
            titulo.setCellValue(nomeTitulo);

            XSSFRow headerRow = sheet.createRow(rowCount++);
            headerRow.setHeightInPoints((short) 20);
            for (int i = 0; i < nomeColunas.size(); i++) {
                XSSFCell cell = headerRow.createCell(i);
                cell.setCellValue(nomeColunas.get(i));
                cell.setCellStyle(headerCellStyle);
            }

            Field[] fields = Arrays.stream(listaDados.get(0).getClass().getDeclaredFields()).filter(field -> field.isAnnotationPresent(ExcelColumn.class)).toArray(Field[]::new);

            for (Object object : listaDados) {
                XSSFRow row = sheet.createRow(rowCount++);
                for (int i = 0; i < fields.length; i++) {
                    XSSFCell cell = row.createCell(i);
                    Object value = this.getFieldByIndex(object, fields, i);
                    cell.setCellStyle(dataCellStyle);
                    if (value instanceof String) cell.setCellValue((String) value);
                    if (value instanceof Integer) cell.setCellValue((Integer) value);
                    if (value instanceof Boolean) cell.setCellValue((Boolean) value);
                    if (value instanceof Enum<?>) cell.setCellValue(((Enum<?>) value).name());
                    if (value instanceof BigDecimal || value instanceof Double)
                        cell.setCellValue(Double.parseDouble(value.toString()));
                }
            }
            for (int i = 0; i < nomeColunas.size(); i++) {
                sheet.autoSizeColumn(i);
            }
            try (ByteArrayOutputStream fileOutputStream = new ByteArrayOutputStream()) {
                workbook.write(fileOutputStream);
                return fileOutputStream.toByteArray();
            }
        } catch (IllegalAccessException e) {
            throw new RuntimeException(e);
        }
    }


    private void validarListaDeDados(List<?> listaDados) throws ApplicationException {
        if (listaDados.isEmpty()) {
            throw ApplicationException.builder()
                    .status(HttpStatus.BAD_REQUEST)
                    .messageKey(MessageKeysEnum.LISTA_DADOS_PARA_PROCESSAMENTO_ESTA_VAZIA)
                    .build();
        }
        Object first = listaDados.getFirst();
        Class<?> aClass = first.getClass();
        if (aClass.getAnnotation(ExcelTable.class) == null) {
            throw ApplicationException.builder()
                    .status(HttpStatus.BAD_REQUEST)
                    .messageKey(MessageKeysEnum.NAO_FOI_POSSIVEL_PROCESSAR_LISTA_DE_DADOS)
                    .build();
        }
    }


    private String extrairNomePlanilhaPorAnotacao(Object object) {
        String nomePlanilha = object.getClass().getAnnotation(ExcelTable.class).name();
        if (nomePlanilha == null || nomePlanilha.isEmpty()) return object.getClass().getSimpleName();
        return nomePlanilha;
    }

    private List<String> extrairColunasPorAnotacoes(Object object) {
        Field[] fields = object.getClass().getDeclaredFields();
        List<String> colunas = new ArrayList<>();
        for (Field field : fields) {
            Optional<ExcelColumn> annotation = Optional.ofNullable(field.getAnnotation(ExcelColumn.class));
            if (annotation.isPresent()) {
                String coluna = annotation.get().name();
                if (coluna.isEmpty()) {
                    colunas.add(field.getName());
                    continue;
                }
                colunas.add(coluna);
            }
        }
        return colunas;
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
