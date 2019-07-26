package com.wiley.poc.excelcompare.model;


import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Color;

public class MarkedPaper {

    private String cellIndex;
    private String expectedAnswer;
    private String submittedAnswer;
    private String expectedFormula;
    private String submittedFormula;
    private String statusMessage;
    private int rowId;
    private int columnId;
    private CellType cellType;
    private String cellForeGroundColor;
    private STATUS status;

    public enum STATUS{
        CORRECT, WRONG, PARTIAL;
    }

    public STATUS getStatus() {
        return status;
    }

    public void setStatus(STATUS status) {
        this.status = status;
    }

    public String getCellForeGroundColor() {
        return cellForeGroundColor;
    }

    public void setCellForeGroundColor(String cellForeGroundColor) {
        this.cellForeGroundColor = cellForeGroundColor;
    }

    public int getRowId() {
        return rowId;
    }

    public void setRowId(int rowId) {
        this.rowId = rowId;
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

    public String getCellIndex() {
        return cellIndex;
    }

    public void setCellIndex(String cellIndex) {
        this.cellIndex = cellIndex;
    }

    public String getExpectedFormula() {
        return expectedFormula;
    }

    public void setExpectedFormula(String expectedFormula) {
        this.expectedFormula = expectedFormula;
    }

    public String getSubmittedFormula() {
        return submittedFormula;
    }

    public void setSubmittedFormula(String submittedFormula) {
        this.submittedFormula = submittedFormula;
    }

    public String getStatusMessage() {
        return statusMessage;
    }

    public void setStatusMessage(String statusMessage) {
        this.statusMessage = statusMessage;
    }

    public String getExpectedAnswer() {
        return expectedAnswer;
    }

    public void setExpectedAnswer(String expectedAnswer) {
        this.expectedAnswer = expectedAnswer;
    }

    public String getSubmittedAnswer() {
        return submittedAnswer;
    }

    public void setSubmittedAnswer(String submittedAnswer) {
        this.submittedAnswer = submittedAnswer;
    }


}
