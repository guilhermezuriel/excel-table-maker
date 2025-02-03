package com.guilhermezuriel.exceltablemaker.dtos;

import com.fasterxml.jackson.annotation.JsonProperty;
import jakarta.validation.constraints.NotNull;
import lombok.*;
import org.hibernate.validator.constraints.Length;

import java.util.AbstractList;


@Builder
public record RequestExcelTable(
     String name,
     Object typeReference,
    @NotNull
    @Length(min = 1)
    @JsonProperty("data")
    AbstractList<?> data,
    StyleExcelTable style){
}
