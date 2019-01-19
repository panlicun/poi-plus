package com.plc.model;

public class MergeColumn {

    private String column;
    private int mergeCount;

    public MergeColumn() {
    }

    public MergeColumn(String column, int mergeCount) {
        this.column = column;
        this.mergeCount = mergeCount;
    }

    public String getColumn() {
        return column;
    }

    public void setColumn(String column) {
        this.column = column;
    }

    public int getMergeCount() {
        return mergeCount;
    }

    public void setMergeCount(int mergeCount) {
        this.mergeCount = mergeCount;
    }
}
