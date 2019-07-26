package com.wiley.poc.excelcompare.service.impl;

import com.wiley.poc.excelcompare.model.CellCounts;
import com.wiley.poc.excelcompare.service.ExcelTableGenerator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.util.ArrayList;
import java.util.HashMap;

@Service
public class ExcelTableGeneratorImpl implements ExcelTableGenerator {
    public HashMap<String, CellCounts> getCellIndexes() {
        HashMap<String, CellCounts> cell_coordinates = new HashMap<String, CellCounts>();
        try {
            FileInputStream excelOne = new FileInputStream(new File("C:\\_0_dev\\projects\\poc-excel-compare\\src\\main\\resources\\files\\assignement_structure.xlsx"));
            Workbook workbook1 = new XSSFWorkbook(excelOne);
            Sheet answer_sheet = workbook1.getSheetAt(1);
            cell_coordinates = rowAndColumnCount(answer_sheet);
        } catch (
                Exception e) {
            e.printStackTrace();
        }
        return cell_coordinates;
    }

    public HashMap<String, CellCounts> rowAndColumnCount(Sheet sheet) {
        HashMap<String, CellCounts> counts = new HashMap<String, CellCounts>();
        int row_count = sheet.getPhysicalNumberOfRows();
        int column_count = sheet.getRow(0).getPhysicalNumberOfCells();
        CellCounts cc = new CellCounts();
        ArrayList<String> arr = new ArrayList<>();
        String [] cell_reference = new String[0];
        Row row = sheet.getRow(0);
        for (Cell cell : row) {
            String ref_string = ((XSSFCell) cell).getReference();
            cell_reference = ref_string.split("(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)");
            arr.add(cell_reference[0]);
        }
        cc.setColumnHeader(arr);
        cc.setRowCount(row_count);
        cc.setColumnCount(column_count);
        counts.put("coordinates", cc);
        System.out.println(counts);
        return counts;
    }

}
