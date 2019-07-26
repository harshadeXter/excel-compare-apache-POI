package com.wiley.poc.excelcompare.service;

import com.wiley.poc.excelcompare.model.CellDetails;
import java.util.HashMap;

public interface GetExcelStructure {
    HashMap<String, CellDetails> getStructure();
}
