package com.jwebcoder.ignite.splitvalidationreport.service.serviceImpl;

import com.jwebcoder.ignite.splitvalidationreport.service.ExcelOperator;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.concurrent.ConcurrentHashMap;

@Service
public class ExcelOperatorImpl implements ExcelOperator {

    /**
     * 存储报表名字，以及相应报表的路径
     */
    private static final Map<String, ArrayList<String>> reportRepository = new ConcurrentHashMap<>();

    @Autowired
    @Qualifier("countryMapping")
    private Map countryMapping;

    @Autowired
    @Qualifier("configInfo")
    private Properties configInfo;

    @Override
    public XSSFSheet getXSSFSheetByName(String country, String sourceFileName) {

        return getXSSFWorkbookByName(country, sourceFileName).getSheetAt(0);

    }

    @Override
    public XSSFWorkbook getXSSFWorkbookByName(String sourceFileName) {

        XSSFWorkbook workbook = null;

        try (FileInputStream fileInputStream = new FileInputStream(new File(sourceFileName))) {

            workbook = new XSSFWorkbook(fileInputStream);
            XSSFSheet sheet = workbook.getSheetAt(0);//创建工作表(Sheet)

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return workbook;
    }


    @Override
    public XSSFWorkbook getXSSFWorkbookByName(String country, String sourceFileName) {
        String outputFileRootPath = configInfo.getProperty("outputFileRootPath");
        int index = -1;
        XSSFWorkbook workbook = null;

        if ((index = reportRepository.get(country).indexOf(outputFileRootPath + File.separatorChar + country + " " + sourceFileName)) != -1) {

            try (FileInputStream fileInputStream = new FileInputStream(new File(String.valueOf(reportRepository.get(country).indexOf(index))))) {

                workbook = new XSSFWorkbook(fileInputStream);
                XSSFSheet sheet = workbook.getSheetAt(0);//创建工作表(Sheet)

            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            }


        } else {

            workbook = new XSSFWorkbook();
            XSSFSheet sheet = workbook.createSheet();

        }

        return workbook;
    }

    @Override
    public List<XSSFRow> readHeader(String country, String sourceFileName) {
        return null;
    }

    @Override
    public void writeHeader(String country, String sourceFileName, List<XSSFRow> rows) {

    }

    @Override
    public List<XSSFRow> readDataBody(String country, String sourceFileName) {
        return null;
    }

    @Override
    public void writeDataBody(String country, String sourceFileName, List<XSSFRow> rows) {

    }
}
