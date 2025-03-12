package com.guilhermezuriel.exceltablemaker.service.excel;


import com.guilhermezuriel.dtos.RequestExcelTable;
import com.guilhermezuriel.dtos.StyleExcelTable;

import java.io.IOException;
import java.util.List;

public interface ExcelService {

    byte[] createExcelByRequest(RequestExcelTable request) throws IOException;
    byte[] createExcelByAnnotatedClassList(List<?> classList, StyleExcelTable style) throws IOException;
}
