package com.guilhermezuriel.exceltablemaker.dtos;

import com.fasterxml.jackson.annotation.JsonProperty;
import jakarta.validation.constraints.NotNull;
import lombok.*;
import org.hibernate.validator.constraints.Length;

import java.util.AbstractList;
import java.util.List;

@Getter
@Builder
public class RequestExcelTable {
    private String name;
    private Object referenceObject;
    @NotNull@Length(min = 1)
    @JsonProperty("data")
    private AbstractList<?> data;
    private StyleExcelTable style;

    public Object getReferenceObject() {
        return referenceObject;
    }

    public List<?> getData() {
        return data;
    }

    public StyleExcelTable getStyle() {
        return style;
    }

    public String getName() {
        return name;
    }
}
