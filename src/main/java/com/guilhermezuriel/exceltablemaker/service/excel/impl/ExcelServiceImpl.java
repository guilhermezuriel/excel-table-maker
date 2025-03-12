package com.guilhermezuriel.exceltablemaker.service.excel.impl;

import com.guilhermezuriel.dtos.RequestExcelTable;
import com.guilhermezuriel.dtos.StyleExcelTable;
import com.guilhermezuriel.exceltablemaker.service.excel.ExcelService;
import com.guilhermezuriel.impl.ExcelGeneratorLocal;
import com.guilhermezuriel.impl.ExcelGeneratorWeb;
import lombok.RequiredArgsConstructor;
import org.springframework.lang.NonNull;
import org.springframework.lang.Nullable;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.AbstractList;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ExcelServiceImpl implements ExcelService {


    private final ExcelGeneratorLocal excelGeneratorLocal = new ExcelGeneratorLocal();
    private final ExcelGeneratorWeb excelGeneratorWeb = new ExcelGeneratorWeb();

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
