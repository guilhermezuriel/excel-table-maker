package com.guilhermezuriel.exceltablemaker.service.excel.impl;

import com.guilhermezuriel.exceltablemaker.service.excel.impl.excelGenerator.impl.ExcelGeneratorCsv;
import com.guilhermezuriel.exceltablemaker.service.excel.impl.excelGenerator.impl.ExcelGeneratorLocal;
import com.guilhermezuriel.exceltablemaker.service.excel.impl.excelGenerator.impl.ExcelGeneratorWeb;
import com.guilhermezuriel.exceltablemaker.service.excel.ExcelService;
import com.guilhermezuriel.exceltablemaker.service.excel.dtos.RequestExcelTable;
import com.guilhermezuriel.exceltablemaker.service.excel.dtos.StyleExcelTable;
import lombok.RequiredArgsConstructor;
import org.springframework.lang.NonNull;
import org.springframework.lang.Nullable;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.AbstractList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {

    private final ExcelGeneratorLocal excelGeneratorLocal;
    private final ExcelGeneratorWeb excelGeneratorWeb;
    private final ExcelGeneratorCsv excelGeneratorCsv;

    @Override
    public byte[] createExcelByCsv(MultipartFile file) {
        if(file.isEmpty()){
            throw new RuntimeException();
        }
        return this.excelGeneratorCsv.generateExcelTable();
    }

    @Override
    public byte[] createExcelByRequest(@NonNull RequestExcelTable request) throws IOException {
        var style = request.style() != null ? request.style() : new StyleExcelTable();
        if(request.name() == null || request.name().isBlank()) {
            return this.excelGeneratorWeb.generateExcelTable(request.data(), style);
        }
        return this.excelGeneratorWeb.generateExcelTable(request.data(), request.name(), style);
    }

    @Override
    public byte[] createExcelByAnnotatedClassList(@NonNull List<?> classList, @Nullable StyleExcelTable style) throws IOException {
         style = style != null ? style : new StyleExcelTable();
        return this.excelGeneratorLocal.generateExcelTable((AbstractList<?>) classList, style);
    }
}
