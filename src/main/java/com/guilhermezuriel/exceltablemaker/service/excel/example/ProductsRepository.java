package com.guilhermezuriel.exceltablemaker.service.excel.example;

import org.springframework.data.jpa.repository.JpaRepository;

public interface ProductsRepository extends JpaRepository<ProductsEntity, Long> {
}
