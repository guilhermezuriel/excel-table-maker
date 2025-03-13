package com.guilhermezuriel.exceltablemaker.service.excel.example;

import com.guilhermezuriel.dtos.StyleExcelTable;
import com.guilhermezuriel.exceltablemaker.service.excel.ExcelService;
import lombok.RequiredArgsConstructor;
import org.springframework.stereotype.Service;

import java.io.IOException;
import java.util.List;

@Service
@RequiredArgsConstructor
public class ProductsService {

    private final ExcelService excelService;
    private final ProductsRepository productsRepository;

    public byte[] generateProductsExcel() throws IOException {
        List<ProductsEntity> productsEntities = this.productsRepository.findAll();
        return this.excelService.createExcelByAnnotatedClassList(productsEntities, new StyleExcelTable());
    }
}
