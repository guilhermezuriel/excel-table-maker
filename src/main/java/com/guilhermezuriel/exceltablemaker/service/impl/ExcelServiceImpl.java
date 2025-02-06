package com.guilhermezuriel.exceltablemaker.service.impl;

import com.guilhermezuriel.exceltablemaker.service.ExcelService;
import com.guilhermezuriel.exceltablemaker.service.dtos.RequestExcelTable;
import com.guilhermezuriel.exceltablemaker.excelGenerator.impl.ExcelGeneratorLocal;
import com.guilhermezuriel.exceltablemaker.excelGenerator.impl.ExcelGeneratorWeb;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.AbstractList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {

    private final ExcelGeneratorLocal excelGeneratorLocal;
    private final ExcelGeneratorWeb excelGeneratorWeb;

    @Override
    public byte[] createExcelByRequest(RequestExcelTable request) throws IOException {
        if(request.name() == null || request.name().isBlank()) {
            return this.excelGeneratorWeb.generateExcelTable(request.data());
        }
        return this.excelGeneratorWeb.generateExcelTable(request.data(), request.name());
    }

    @Override
    public byte[] createExcelByAnnotatedClassList(List<?> classList) throws IOException {
        return this.excelGeneratorLocal.generateExcelTable((AbstractList<?>) classList);
    }
}
