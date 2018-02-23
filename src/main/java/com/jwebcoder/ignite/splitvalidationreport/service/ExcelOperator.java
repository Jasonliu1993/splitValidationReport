package com.jwebcoder.ignite.splitvalidationreport.service;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public interface ExcelOperator {

    XSSFSheet getXSSFSheetByName(String country, String sourceFileName);

    XSSFWorkbook getXSSFWorkbookByName(String sourceFileName);

    XSSFWorkbook getXSSFWorkbookByName(String country, String sourceFileName);

    List<XSSFRow> readHeader(String country, String sourceFileName);
    void writeHeader(String country, String sourceFileName, List<XSSFRow> rows);

    List<XSSFRow> readDataBody(String country, String sourceFileName);
    void writeDataBody(String country, String sourceFileName, List<XSSFRow> rows);

}
