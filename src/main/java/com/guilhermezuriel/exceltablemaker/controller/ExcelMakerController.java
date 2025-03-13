package com.guilhermezuriel.exceltablemaker.controller;

import com.guilhermezuriel.dtos.RequestExcelTable;
import com.guilhermezuriel.exceltablemaker.service.excel.ExcelService;
import com.guilhermezuriel.exceltablemaker.service.excel.example.ProductsService;
import lombok.RequiredArgsConstructor;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import java.io.IOException;

@RestController
@RequestMapping("v1/excel")
@RequiredArgsConstructor
public class ExcelMakerController {

    private final ExcelService excelService;
    private final ProductsService productsService;



    @PostMapping
    public ResponseEntity<byte[]> createExcel(@RequestBody RequestExcelTable requestExcelTable) throws IOException {
        var inputStream = this.excelService.createExcelByRequest(requestExcelTable);
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", requestExcelTable.name()+".xlsx");
        return ResponseEntity.ok().headers(headers).body(inputStream);
    }

    @PostMapping("/local-example")
    public ResponseEntity<byte[]> createLocalProcessingExample() throws IOException {
        var inputStream = this.productsService.generateProductsExcel();
        HttpHeaders headers = new HttpHeaders();
        headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
        headers.setContentDispositionFormData("attachment", "products.xlsx");
        return ResponseEntity.ok().headers(headers).body(inputStream);
    }


}
