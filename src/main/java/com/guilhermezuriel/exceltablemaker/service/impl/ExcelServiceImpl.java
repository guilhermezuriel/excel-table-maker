package com.guilhermezuriel.exceltablemaker.service.impl;

import com.guilhermezuriel.exceltablemaker.service.ExcelService;
import com.guilhermezuriel.exceltablemaker.service.dtos.RequestExcelTable;
import com.guilhermezuriel.exceltablemaker.excelGenerator.impl.ExcelGeneratorLocal;
import com.guilhermezuriel.exceltablemaker.excelGenerator.impl.ExcelGeneratorWeb;
import com.guilhermezuriel.exceltablemaker.service.dtos.StyleExcelTable;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.AbstractList;
import java.util.List;
import java.util.Objects;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {

    private final ExcelGeneratorLocal excelGeneratorLocal;
    private final ExcelGeneratorWeb excelGeneratorWeb;

    @Override
    public byte[] createExcelByRequest(RequestExcelTable request) throws IOException {
        var style = request.style() != null ? request.style() : new StyleExcelTable();
        if(request.name() == null || request.name().isBlank()) {
            return this.excelGeneratorWeb.generateExcelTable(request.data(), style);
        }
        return this.excelGeneratorWeb.generateExcelTable(request.data(), request.name(), style);
    }

    @Override
    public byte[] createExcelByAnnotatedClassList(List<?> classList, StyleExcelTable style) throws IOException {
         style = style != null ? style : new StyleExcelTable();
        return this.excelGeneratorLocal.generateExcelTable((AbstractList<?>) classList, style);
    }
}
