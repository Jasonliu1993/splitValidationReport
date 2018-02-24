package com.jwebcoder.ignite.splitvalidationreport.service;

import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.List;

public interface ExcelOperator {

    XSSFWorkbook getSourceWorkbook(String sourceFileKey);
    XSSFWorkbook getOutputWorkbook(String country, String sourceFileKey);
    boolean insertRow(XSSFRow source, XSSFRow target);
    List<XSSFRow> readHeader(String sourceFileKey);
    void writeHeader(String country, String sourceFileKey, List<XSSFRow> rows);
    List<XSSFRow> readDataBody(String sourceFileKey,int pageIndex, int pageCount);
    void writeDataBody(String country, String sourceFileKey, List<XSSFRow> rows);
    boolean saveAllWorkbook();
    boolean saveSingleWorkbook(String country, String sourceFileKey);
    void saveWorkbook(String fullpath) throws Exception;

}
