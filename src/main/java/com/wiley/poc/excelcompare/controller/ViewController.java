package com.wiley.poc.excelcompare.controller;

import com.wiley.poc.excelcompare.service.ReadExcelFile;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.ResponseBody;

@Controller
public class ViewController {

    @Autowired
    ReadExcelFile readExcelFile;

    @GetMapping("/excel/test")
    public String resultHTML(Model model){
        model.addAttribute("myJson", readExcelFile.compareExcels());
        return "resultTable";
    }
}
