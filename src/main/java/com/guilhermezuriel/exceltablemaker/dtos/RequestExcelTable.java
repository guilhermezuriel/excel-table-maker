package com.guilhermezuriel.exceltablemaker.dtos;

import jakarta.validation.constraints.NotNull;
import lombok.*;
import org.hibernate.validator.constraints.Length;

import java.util.List;

@Data
@Builder
public class RequestExcelTable {
    private String name;
    private Object referenceObject;
    @NotNull@Length(min = 1)
    private List<?> data;
    private StyleExcelTable style;
}
