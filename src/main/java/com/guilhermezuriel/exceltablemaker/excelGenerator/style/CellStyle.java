package com.guilhermezuriel.exceltablemaker.excelGenerator.style;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.VerticalAlignment;

public record CellStyle(
        short fontHeight,
        IndexedColors fontColor,
        IndexedColors backgroundColor,
        FillPatternType fillPattern,
        boolean isBold,
        boolean wrapText,
        VerticalAlignment verticalAlignment,
        HorizontalAlignment horizontalAlignment,
        BorderCellStyle borderProps
) {
    public static CellStyle defaultTitle() {
        return new CellStyle((short) 13,
                IndexedColors.WHITE,
                IndexedColors.BLACK,
                FillPatternType.FINE_DOTS,
                true,
                false,
                VerticalAlignment.CENTER,
                HorizontalAlignment.CENTER_SELECTION,
                BorderCellStyle.title());
    }

    public static CellStyle defaultHeader() {
        return new CellStyle(
                (short) 13,
                IndexedColors.WHITE,
                IndexedColors.BLACK,
                FillPatternType.FINE_DOTS,
                false,
                false,
                VerticalAlignment.CENTER,
                HorizontalAlignment.CENTER,
                BorderCellStyle.header());
    }

    public static CellStyle defaultData() {
        return new CellStyle(
                (short) 12,
                IndexedColors.BLACK,
                IndexedColors.WHITE,
                FillPatternType.NO_FILL,
                false,
                true,
                VerticalAlignment.CENTER,
                HorizontalAlignment.CENTER,
                BorderCellStyle.data()
        );
    }
}
