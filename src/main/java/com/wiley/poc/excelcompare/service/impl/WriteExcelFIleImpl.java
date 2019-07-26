package com.wiley.poc.excelcompare.service.impl;

import com.wiley.poc.excelcompare.controller.ExcelController;
import com.wiley.poc.excelcompare.model.CellCounts;
import com.wiley.poc.excelcompare.service.ExcelTableGenerator;
import com.wiley.poc.excelcompare.service.WriteExcelFile;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.HashMap;

@Service
public class WriteExcelFIleImpl implements WriteExcelFile {
    public void writeExcel() {

    }
}
