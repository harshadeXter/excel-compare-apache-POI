package com.wiley.poc.excelcompare.model;

import java.lang.reflect.Array;
import java.util.ArrayList;

public class CellCounts {
    private int rowCount;
    private int columnCount;
    private ArrayList<String > columnHeader;

    public ArrayList<String> getColumnHeader() {
        return columnHeader;
    }

    public void setColumnHeader(ArrayList<String> columnHeader) {
        this.columnHeader = columnHeader;
    }

    public int getRowCount() {
        return rowCount;
    }

    public void setRowCount(int rowCount) {
        this.rowCount = rowCount;
    }

    public int getColumnCount() {
        return columnCount;
    }

    public void setColumnCount(int columnCount) {
        this.columnCount = columnCount;
    }
}
