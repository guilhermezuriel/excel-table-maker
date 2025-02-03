package com.guilhermezuriel.exceltablemaker.controller;

import com.guilhermezuriel.exceltablemaker.dtos.RequestExcelTable;
import com.guilhermezuriel.exceltablemaker.services.ExcelGeneratorWeb;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@RequestMapping("v1")
public class ExcelMakerController {

    private  ExcelGeneratorWeb excelGeneratorWeb;

    public ExcelMakerController(ExcelGeneratorWeb excelGeneratorWeb){
        this.excelGeneratorWeb = excelGeneratorWeb;
    }

    @PostMapping("excel")
    public ResponseEntity<Object> createExcel(@RequestBody RequestExcelTable requestExcelTable) throws IOException {
        this.excelGeneratorWeb.criarPlanilhaExcel(requestExcelTable);
        return ResponseEntity.ok().build();
    }
}
