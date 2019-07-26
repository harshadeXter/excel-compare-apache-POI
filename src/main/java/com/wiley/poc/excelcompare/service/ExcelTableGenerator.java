package com.wiley.poc.excelcompare.service;

import com.wiley.poc.excelcompare.model.CellCounts;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.HashMap;

public interface ExcelTableGenerator {
    HashMap<String, CellCounts> getCellIndexes();

    HashMap<String, CellCounts> rowAndColumnCount(Sheet sheet);
}
