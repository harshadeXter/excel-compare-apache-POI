package com.wiley.poc.excelcompare.service.impl;


import com.wiley.poc.excelcompare.model.CellDetails;
import com.wiley.poc.excelcompare.service.GetExcelStructure;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;

@Service
public class GetExcelStructureImpl implements GetExcelStructure {
    public HashMap<String, CellDetails> getStructure() {
        HashMap<String, CellDetails> cell_details = new HashMap<String, CellDetails>();
        try {
            /*Read the two excel files from local directory*/
            FileInputStream excelOne = new FileInputStream(new File("C:\\_0_dev\\projects\\poc-excel-compare\\src\\main\\resources\\files\\assignement_structure.xlsx"));
            /*Create workbook instance using apache POI which refers to excel files*/
            Workbook workbook1 = new XSSFWorkbook(excelOne);
            /*Select the first sheet in excel file*/
            Sheet answer_sheet = workbook1.getSheetAt(1);
            CellDetails cd;
            for (Row row : answer_sheet) {
                for (Cell cell : row) {
                    cd = new CellDetails();
                    String cell_info = ((XSSFCell) cell).getStringCellValue();
                    String cell_reference = ((XSSFCell) cell).getReference();
                    String[] cell_index_ = cell_reference.split("(?<=\\D)(?=\\d)|(?<=\\d)(?=\\D)");
                    cd.setCellRef(cell_reference);
                    cd.setAnswerValue(cell_info);
                    cd.setColumnIndex(cell_index_[0]);
                    cd.setRowIndex(Integer.parseInt(cell_index_[1]));
                    cell_details.put(cell_reference, cd);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return cell_details;
    }
}
