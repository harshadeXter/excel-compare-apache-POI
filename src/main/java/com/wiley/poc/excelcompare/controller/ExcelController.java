package com.wiley.poc.excelcompare.controller;

import com.wiley.poc.excelcompare.model.MarkedPaper;
import com.wiley.poc.excelcompare.service.ExcelTableGenerator;
import com.wiley.poc.excelcompare.service.GetExcelStructure;
import com.wiley.poc.excelcompare.service.ReadExcelFile;
import com.wiley.poc.excelcompare.service.WriteExcelFile;
import org.json.JSONObject;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RestController;
import java.util.HashMap;

@RestController
public class ExcelController {
    @Autowired
    ReadExcelFile readExcelFile;
    @Autowired
    GetExcelStructure getExcelStructure;
    @Autowired
    ExcelTableGenerator excelTableGenerator;
    @Autowired
    WriteExcelFile writeExcelFile;

    @GetMapping("/results")
    public HashMap<Double, MarkedPaper> compare() {
        return readExcelFile.compareExcels();
    }

    @CrossOrigin(origins = "http://localhost:8081")
    @GetMapping("/table/excel")
    public String showStructure() {
        JSONObject data = new JSONObject();
        data.put("structure",excelTableGenerator.getCellIndexes());
        data.put("excelData",getExcelStructure.getStructure());
        return data.toString();
    }

/*    @CrossOrigin(origins = "http://localhost:8082")
    @GetMapping("/table")
    public HashMap<String, CellCounts> cellIndexes() {
        return excelTableGenerator.getCellIndexes();
    }*/


    @GetMapping("/table")
    public void writeExcelDoc(){
        writeExcelFile.writeExcel();
    }
}
