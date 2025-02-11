package com.guilhermezuriel.exceltablemaker.service.excel.dtos;

import com.guilhermezuriel.exceltablemaker.service.excel.impl.excelGenerator.style.CellStyle;
import lombok.Getter;

import java.util.Objects;

@Getter
public class StyleExcelTable {
    private CellStyle title;
    private CellStyle header;
    private CellStyle data;

    public StyleExcelTable() {
        this.title = CellStyle.defaultTitle();
        this.header = CellStyle.defaultHeader();
        this.data = CellStyle.defaultData();
    }

    public StyleExcelTable(CellStyle title, CellStyle header, CellStyle data) {
        if (Objects.isNull(title)) this.title = CellStyle.defaultTitle();
        if (Objects.isNull(header)) this.header = CellStyle.defaultHeader();
        if (Objects.isNull(data)) this.data = CellStyle.defaultData();
    }
}
