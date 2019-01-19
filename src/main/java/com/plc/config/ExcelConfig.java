package com.plc.config;

/**
 * Created by Administrator on 2018/12/30.
 */
public class ExcelConfig {

    //从哪行开始读取数据
    private int readStartRow = 0;
    private int readTitleRow = 1;


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
}
