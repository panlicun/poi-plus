package com.plc.config;

import com.plc.model.ExcelCustomData;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;

/**
 * Created by Administrator on 2018/12/30.
 */
public class ExcelConfig {

    //从哪行开始读取数据
    private int readStartRow = 0;
    private int readTitleRow = 0;

    //从哪行开始写入
    private int writeStartRow = 0;
    private int writeStartCol = 0;

    //自定义单元格内容
    private List<ExcelCustomData> excelCustomDatas;


    public int getReadStartRow() {
        return readStartRow;
    }

    public void setReadStartRow(int readStartRow) {
        this.readStartRow = readStartRow;
    }

    public int getReadTitleRow() {
        return readTitleRow;
    }

    public void setReadTitleRow(int readTitleRow) {
        this.readTitleRow = readTitleRow;
    }

    public int getWriteStartRow() {
        return writeStartRow;
    }

    public void setWriteStartRow(int writeStartRow) {
        this.writeStartRow = writeStartRow;
    }

    public int getWriteStartCol() {
        return writeStartCol;
    }

    public void setWriteStartCol(int writeStartCol) {
        this.writeStartCol = writeStartCol;
    }

    public List<ExcelCustomData> getExcelCustomDatas() {
        return excelCustomDatas;
    }

    public void setExcelCustomDatas(List<ExcelCustomData> excelCustomDatas) {
        this.excelCustomDatas = excelCustomDatas;
    }
}
