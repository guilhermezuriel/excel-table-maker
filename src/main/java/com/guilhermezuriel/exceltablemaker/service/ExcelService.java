package com.guilhermezuriel.exceltablemaker.service;

import com.guilhermezuriel.exceltablemaker.service.dtos.RequestExcelTable;

import java.io.IOException;
import java.util.List;

public interface ExcelService {

    byte[] createExcelByRequest(RequestExcelTable request) throws IOException;
    byte[] createExcelByAnnotatedClassList(List<?> classList) throws IOException;
}
