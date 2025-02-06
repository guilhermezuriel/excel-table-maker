package com.guilhermezuriel.exceltablemaker.controller;

import com.guilhermezuriel.exceltablemaker.service.ExcelService;
import com.guilhermezuriel.exceltablemaker.service.dtos.RequestExcelTable;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@RequestMapping("v1")
public class ExcelMakerController {

    private ExcelService excelService;

    public ExcelMakerController(ExcelService excelGeneratorWeb){
        this.excelService = excelGeneratorWeb;
    }

    @PostMapping("excel")
    public ResponseEntity<byte[]> createExcel(@RequestBody RequestExcelTable requestExcelTable) throws IOException {
        var inputStream = this.excelService.createExcelByRequest(requestExcelTable);
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", requestExcelTable.name()+".xlsx");
        return ResponseEntity.ok().headers(headers).body(inputStream);
    }
}
