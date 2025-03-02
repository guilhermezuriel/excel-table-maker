package com.guilhermezuriel.exceltablemaker.service.excel;

import com.guilhermezuriel.exceltablemaker.service.excel.dtos.RequestExcelTable;
import com.guilhermezuriel.exceltablemaker.service.excel.dtos.StyleExcelTable;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

public interface ExcelService {
    byte[] createExcelByCsv(MultipartFile file);
    byte[] createExcelByRequest(RequestExcelTable request) throws IOException;
    byte[] createExcelByAnnotatedClassList(List<?> classList, StyleExcelTable style) throws IOException;
}
