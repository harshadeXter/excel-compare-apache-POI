package com.wiley.poc.excelcompare.model;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;

public class CellDetails {

    private String cellRef;
    private String answerFormulae;
    private String answerValue;
    private String columnIndex;
    private int rowIndex;
    private int columnId;
    private CellType cellType;
    private CellStyle cellStyle;
    private String foreGroundColor;
    private Color cellColor;

    public Color getCellColor() {
        return cellColor;
    }

    public void setCellColor(Color cellColor) {
        this.cellColor = cellColor;
    }

    public String getForeGroundColor() {
        return foreGroundColor;
    }

    public void setForeGroundColor(String foreGroundColor) {
        this.foreGroundColor = foreGroundColor;
    }

    public int getColumnId() {
        return columnId;
    }

    public void setColumnId(int columnId) {
        this.columnId = columnId;
    }

    public CellType getCellType() {
        return cellType;
    }

    public void setCellType(CellType cellType) {
        this.cellType = cellType;
    }

    public CellStyle getCellStyle() {
        return cellStyle;
    }

    public void setCellStyle(CellStyle cellStyle) {
        this.cellStyle = cellStyle;
    }

    public String getColumnIndex() {
        return columnIndex;
    }

    public void setColumnIndex(String columnIndex) {
        this.columnIndex = columnIndex;
    }

    public int getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(int rowIndex) {
        this.rowIndex = rowIndex;
    }

    public String getCellRef() {
        return cellRef;
    }

    public void setCellRef(String cellRef) {
        this.cellRef = cellRef;
    }

    public String getAnswerFormulae() {
        return answerFormulae;
    }

    public void setAnswerFormulae(String answerFormulae) {
        this.answerFormulae = answerFormulae;
    }

    public String getAnswerValue() {
        return answerValue;
    }

    public void setAnswerValue(String answerValue) {
        this.answerValue = answerValue;
    }
}
