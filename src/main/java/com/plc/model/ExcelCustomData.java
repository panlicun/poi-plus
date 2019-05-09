package com.plc.model;

import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelCustomData {

    private String excelText;

    private CellRangeAddress cellRangeAddress;


    public ExcelCustomData(String excelText, CellRangeAddress cellRangeAddress) {
        this.excelText = excelText;
        this.cellRangeAddress = cellRangeAddress;
    }

    public String getExcelText() {
        return excelText;
    }

    public void setExcelText(String excelText) {
        this.excelText = excelText;
    }

    public CellRangeAddress getCellRangeAddress() {
        return cellRangeAddress;
    }

    public void setCellRangeAddress(CellRangeAddress cellRangeAddress) {
        this.cellRangeAddress = cellRangeAddress;
    }

}
