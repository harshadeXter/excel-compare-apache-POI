package com.wiley.poc.excelcompare.service.impl;


import com.wiley.poc.excelcompare.model.CellDetails;
import com.wiley.poc.excelcompare.model.MarkedPaper;
import com.wiley.poc.excelcompare.service.ReadExcelFile;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.*;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellBorder;
import org.springframework.stereotype.Service;


import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.HashMap;


@Service
public class ReadExcelFileImpl implements ReadExcelFile {

    public HashMap<Double, MarkedPaper> compareExcels() {
        HashMap results = new HashMap<Double, MarkedPaper>();
        try {
            /*Read the two excel files from local directory*/
            FileInputStream excelOne = new FileInputStream(new File("C:\\_0_dev\\projects\\poc-excel-compare\\src\\main\\resources\\files\\instructor_sheet\\answer_sheet.xlsx"));
            FileInputStream excelTwo = new FileInputStream(new File("C:\\_0_dev\\projects\\poc-excel-compare\\src\\main\\resources\\files\\student_sheet\\answer_sheet.xlsx"));
            /*Create workbook instance using apache POI which refers to excel files*/
            Workbook workbook1 = new XSSFWorkbook(excelOne);
            Workbook workbook2 = new XSSFWorkbook(excelTwo);

            /*Select the first sheet in excel file*/
            Sheet answer_sheet = workbook1.getSheetAt(0);
            Sheet student_sheet = workbook2.getSheetAt(0);
            HashMap answer_sheet_map = getInstructorAnswers(answer_sheet);
            HashMap student_sheet_map = getStudentAnswers(student_sheet);


            results = validateAnswers(answer_sheet_map, student_sheet_map);
            markAnswerSheet(results);

        } catch (Exception e) {
            e.printStackTrace();
        }
        return results;
    }

    /*Method to store Instructor answer sheet*/
    public HashMap<String, Object> getInstructorAnswers(Sheet sheet) {
        HashMap<String, Object> cell_details = new HashMap<String, Object>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                CellDetails cd = new CellDetails();
                String ref_string = ((XSSFCell) cell).getReference();
                CellStyle cs = cell.getCellStyle();
                Color color = cs.getFillForegroundColorColor();
                if (color != null) {
                    if (color instanceof XSSFColor) {
                        if (((XSSFColor) color).getARGBHex().equals("FFFFFF99")) {
                            CellType cell_type;
                            if (cell != null) {
                                cell_type = cell.getCellType();
                                cd.setCellRef(ref_string);
                                if (cell_type == CellType.FORMULA && cell_type != CellType.BLANK) {
                                    cd.setAnswerFormulae(cell.getCellFormula());
                                } else if (cell_type == CellType.NUMERIC && cell_type != CellType.BLANK) {
                                    String cell_value = ((XSSFCell) cell).getRawValue();
                                    cd.setAnswerValue(cell_value);
                                }
                            } else {
                                cd.setAnswerFormulae(null);
                                cd.setAnswerValue(null);
                            }
                            cell_details.put(ref_string, cd);
                        }
                    }
                }
            }
        }
        return cell_details;
    }

    /*Method to store Student answer sheet*/
    public HashMap<String, Object> getStudentAnswers(Sheet sheet) {
        HashMap<String, Object> cell_details = new HashMap<String, Object>();
        for (Row row : sheet) {
            for (Cell cell : row) {
                CellDetails cd = new CellDetails();
                String ref_string = ((XSSFCell) cell).getReference();
                CellStyle cs = cell.getCellStyle();
                Color color = cs.getFillForegroundColorColor();
                int rowId = cell.getRowIndex();
                int columnId = cell.getColumnIndex();
                cd.setRowIndex(rowId);
                cd.setColumnId(columnId);
                if (color != null) {
                    if (color instanceof XSSFColor) {
                        String color_code = ((XSSFColor) color).getARGBHex();
                        cd.setForeGroundColor(color_code);
                        if (((XSSFColor) color).getARGBHex().equals("FFFFFF99")) {
                            CellType cell_type;
                            if (cell != null) {
                                cell_type = cell.getCellType();
                                cd.setCellRef(ref_string);
                                if (cell_type == CellType.FORMULA && cell_type != CellType.BLANK) {
                                    cd.setAnswerFormulae(cell.getCellFormula());
                                } else if (cell_type == CellType.NUMERIC && cell_type != CellType.BLANK) {
                                    String cell_value = ((XSSFCell) cell).getRawValue();
                                    cd.setAnswerValue(cell_value);
                                }
                            } else {
                                cd.setAnswerFormulae(null);
                                cd.setAnswerValue(null);
                            }
                            cell_details.put(ref_string, cd);
                        }
                    }
                }
            }
        }
        return cell_details;
    }

    /*Compare hash map to validate the answers*/
    public HashMap<String, MarkedPaper> validateAnswers(HashMap instructorSheetMap, HashMap studentSheetMap) {
        HashMap<String, MarkedPaper> markedPaper = new HashMap<String, MarkedPaper>();
        MarkedPaper paper;

        for (Object key : instructorSheetMap.keySet()) {
            String k = (String) key;
            CellDetails student_answers = (CellDetails) studentSheetMap.get(key);
            CellDetails instructor_answers = (CellDetails) instructorSheetMap.get(key);

            paper = new MarkedPaper();
            paper.setRowId(student_answers.getRowIndex());
            paper.setColumnId(student_answers.getColumnId());
            paper.setCellForeGroundColor(student_answers.getForeGroundColor());
            paper.setCellType(student_answers.getCellType());
            if (student_answers.getAnswerFormulae() != null && student_answers.getAnswerFormulae().equals(instructor_answers.getAnswerFormulae())) {
                paper.setCellIndex(k);
                paper.setSubmittedFormula(student_answers.getAnswerFormulae());
                paper.setExpectedFormula(instructor_answers.getAnswerFormulae());
                paper.setSubmittedAnswer(student_answers.getAnswerValue());
                paper.setExpectedAnswer(instructor_answers.getAnswerValue());
                paper.setStatusMessage("Correct Answer given");
                paper.setStatus(MarkedPaper.STATUS.CORRECT);
            } else if (student_answers.getAnswerValue() != null && student_answers.getAnswerValue().equals(instructor_answers.getAnswerValue())) {
                paper.setCellIndex(k);
                paper.setSubmittedAnswer(student_answers.getAnswerValue());
                paper.setExpectedAnswer(instructor_answers.getAnswerValue());
                paper.setSubmittedFormula(student_answers.getAnswerFormulae());
                paper.setExpectedFormula(instructor_answers.getAnswerFormulae());
                paper.setStatusMessage("Partially Correct Answer given");
                paper.setStatus(MarkedPaper.STATUS.PARTIAL);
            } else {
                paper.setCellIndex(k);
                paper.setSubmittedAnswer(student_answers.getAnswerValue());
                paper.setExpectedAnswer(instructor_answers.getAnswerValue());
                paper.setSubmittedFormula(student_answers.getAnswerFormulae());
                paper.setExpectedFormula(instructor_answers.getAnswerFormulae());
                paper.setStatusMessage("Incorrect Answer given");
                paper.setStatus(MarkedPaper.STATUS.WRONG);
            }
            markedPaper.put(k, paper);
        }
        return markedPaper;
    }

    public void markAnswerSheet(HashMap results) {
        for (Object key : results.keySet()) {
            MarkedPaper cell_info = (MarkedPaper) results.get(key);
            int rowIndex = cell_info.getRowId();
            int columnIndex = cell_info.getColumnId();
            String feedback_statement = cell_info.getStatusMessage();
            try {
                FileInputStream openFile = new FileInputStream(new File("C:\\_0_dev\\projects\\poc-excel-compare\\src\\main\\resources\\files\\student_sheet\\answer_sheet.xlsx"));
                XSSFWorkbook workbook = new XSSFWorkbook(openFile);
                XSSFSheet sheetName = workbook.getSheetAt(0);

                Cell cell = sheetName.getRow(rowIndex).getCell(columnIndex);
                /*Set comments to excel cells*/
                CreationHelper factory = workbook.getCreationHelper();
                Drawing drawing = sheetName.createDrawingPatriarch();
                ClientAnchor anchor = factory.createClientAnchor();
                Comment comment = drawing.createCellComment(anchor);
                RichTextString richy = factory.createRichTextString(feedback_statement);
                comment.setString(richy);
                cell.setCellComment(comment);

                CellStyle cell_style_red = workbook.createCellStyle();
                cell_style_red.setFillForegroundColor(IndexedColors.ROSE.getIndex());
                cell_style_red.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                XSSFFont font_red = workbook.createFont();
                font_red.setColor(IndexedColors.ROSE.getIndex());
                cell_style_red.setFont(font_red);
                cell_style_red.setBorderBottom(BorderStyle.THIN);
                cell_style_red.setBottomBorderColor(IndexedColors.ROSE.getIndex());
                cell_style_red.setBorderTop(BorderStyle.THIN);
                cell_style_red.setTopBorderColor(IndexedColors.ROSE.getIndex());

                CellStyle cell_style_green = workbook.createCellStyle();
                cell_style_green.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                cell_style_green.setFillPattern(FillPatternType.FINE_DOTS);
                cell_style_green.setFillBackgroundColor(IndexedColors.LIGHT_GREEN.getIndex());
                XSSFFont font_green = workbook.createFont();
                font_green.setColor(IndexedColors.GREEN.getIndex());
                cell_style_green.setFont(font_green);
                cell_style_green.setBorderBottom(BorderStyle.THIN);
                cell_style_green.setBottomBorderColor(IndexedColors.GREEN.getIndex());
                cell_style_green.setBorderTop(BorderStyle.THIN);
                cell_style_green.setTopBorderColor(IndexedColors.GREEN.getIndex());

                CellStyle cell_style_blue = workbook.createCellStyle();
                cell_style_blue.setFillForegroundColor(IndexedColors.SKY_BLUE.getIndex());
                cell_style_blue.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cell_style_blue.setFillBackgroundColor(IndexedColors.SKY_BLUE.getIndex());
                XSSFFont font_blue = workbook.createFont();
                font_blue.setColor(IndexedColors.LIGHT_BLUE.getIndex());
                cell_style_blue.setFont(font_blue);
                cell_style_blue.setBorderBottom(BorderStyle.THIN);
                cell_style_blue.setBottomBorderColor(IndexedColors.LIGHT_BLUE.getIndex());
                cell_style_blue.setBorderTop(BorderStyle.THIN);
                cell_style_blue.setTopBorderColor(IndexedColors.LIGHT_BLUE.getIndex());

                if (cell_info.getStatus() == MarkedPaper.STATUS.CORRECT) {
                    /*Image*/
                    FileInputStream stream = new FileInputStream("C:\\\\_0_dev\\\\projects\\\\poc-excel-compare\\\\src\\\\main\\\\resources\\\\images\\\\tick.png");
                    anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
                    int tickIndex = workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_PNG);
                    anchor.setCol1(columnIndex);
                    anchor.setRow1(rowIndex);
                    anchor.setRow2(rowIndex);
                    anchor.setCol2(rowIndex-1);
                    Picture pic = drawing.createPicture(anchor,tickIndex);
                    pic.resize();
                    cell.setCellStyle(cell_style_green);
                } else if (cell_info.getStatus() == MarkedPaper.STATUS.PARTIAL) {
                    cell.setCellStyle(cell_style_blue);
                } else if (cell_info.getStatus() == MarkedPaper.STATUS.WRONG) {
                    /*Image*/
                    FileInputStream stream = new FileInputStream("C:\\\\_0_dev\\\\projects\\\\poc-excel-compare\\\\src\\\\main\\\\resources\\\\images\\\\cross.png");
                    anchor.setAnchorType(ClientAnchor.AnchorType.DONT_MOVE_AND_RESIZE);
                    int tickIndex = workbook.addPicture(IOUtils.toByteArray(stream), Workbook.PICTURE_TYPE_PNG);
                    anchor.setCol1(columnIndex);
                    anchor.setRow1(rowIndex);
                    anchor.setRow2(rowIndex);
                    anchor.setCol2(rowIndex-1);
                    Picture pic = drawing.createPicture(anchor,tickIndex);
                    pic.resize();
                    cell.setCellStyle(cell_style_red);
                }
                try (FileOutputStream outputFile = new FileOutputStream("C:\\_0_dev\\projects\\poc-excel-compare\\src\\main\\resources\\files\\student_sheet\\answer_sheet.xlsx")) {
                    workbook.write(outputFile);
                }
                openFile.close();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }
}
