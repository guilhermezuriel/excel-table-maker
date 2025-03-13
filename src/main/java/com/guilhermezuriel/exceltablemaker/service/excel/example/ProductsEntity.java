package com.guilhermezuriel.exceltablemaker.service.excel.example;

import com.guilhermezuriel.annotations.ExcelColumn;
import com.guilhermezuriel.annotations.ExcelTable;
import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.RequiredArgsConstructor;
import org.hibernate.annotations.Type;

import java.math.BigDecimal;
import java.util.UUID;

@ExcelTable(name = "Stored Products")
@Entity
@Table( name = "products", schema = "example")
@RequiredArgsConstructor
public class ProductsEntity {

    @Id
    @ExcelColumn
    private Long id;

    @ExcelColumn
    @Column(name = "price")
    private BigDecimal price;

    @ExcelColumn(name = "name")
    @Column(name = "product_name")
    private String name;

    @ExcelColumn
    @Column(name = "stock")
    private Integer stock;

    @Column(name = "credential")
    private UUID credential;


}
