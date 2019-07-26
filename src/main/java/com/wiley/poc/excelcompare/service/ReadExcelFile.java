package com.wiley.poc.excelcompare.service;

import com.wiley.poc.excelcompare.model.MarkedPaper;
import org.apache.poi.ss.usermodel.Sheet;


import java.util.HashMap;
import java.util.Map;

public interface ReadExcelFile {
    HashMap<Double, MarkedPaper> compareExcels();

    HashMap<String, Object> getInstructorAnswers(Sheet sheet);

    HashMap<String, Object> getStudentAnswers(Sheet sheet);

    HashMap<String, MarkedPaper> validateAnswers(HashMap map1, HashMap map2);
}
