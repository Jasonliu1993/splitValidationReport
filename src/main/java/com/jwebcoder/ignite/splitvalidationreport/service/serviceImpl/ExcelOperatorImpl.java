package com.jwebcoder.ignite.splitvalidationreport.service.serviceImpl;

import com.jwebcoder.ignite.splitvalidationreport.service.ExcelOperator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.stereotype.Service;

import java.io.*;
import java.util.*;
import java.util.concurrent.ConcurrentHashMap;


@Service
public class ExcelOperatorImpl implements ExcelOperator {

    @Autowired
    @Qualifier("configInfo")
    private Properties configInfo;

    @Autowired
    @Qualifier("sourceFilePropertiesPath")
    private Properties sourceFilePropertiesPath;

    @Autowired
    @Qualifier("countryMapping")
    private Map<String, String> countryMapping;

    @Autowired
    @Qualifier("headerCountMapping")
    private Map<String, Integer> headerCountMapping;

    /**
     * String 是fullPath
     * XSSFWorkbook 对应的workbook
     */
    private static ConcurrentHashMap<String, XSSFWorkbook> workbook = new ConcurrentHashMap<>();

    /**
     * 获取源文件的workbook
     *
     * @param sourceFileKey 是filePath的key
     * @return 获取源文件的workbook
     */
    @Override
    public XSSFWorkbook getSourceWorkbook(String sourceFileKey) {
        String fullPath = sourceFilePropertiesPath.getProperty(sourceFileKey);

        if (workbook.containsKey(fullPath))
            return workbook.get(fullPath);

        FileInputStream fileInputStream = null;
        XSSFWorkbook workbook = null;
        try {
            fileInputStream = new FileInputStream(new File(fullPath));
            workbook = new XSSFWorkbook(fileInputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return workbook;
    }

    /**
     * @param country       如国家或者是No group
     * @param sourceFileKey 是filePath的key
     * @return 获取输出文件的
     */
    @Override
    public XSSFWorkbook getOutputWorkbook(String country, String sourceFileKey) {
        String fullPath = getFullPath(country, sourceFileKey);

        if (workbook.containsKey(fullPath))
            return workbook.get(fullPath);

        FileInputStream fileInputStream = null;
        XSSFWorkbook workbook = null;
        try {
            fileInputStream = new FileInputStream(new File(fullPath));
            workbook = new XSSFWorkbook(fileInputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                fileInputStream.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return workbook;
    }

    @Override
    public boolean insertRow(XSSFRow source, XSSFRow target) {
        try {
            Iterator<Cell> iterator = source.iterator();
            int index = 0;
            while (iterator.hasNext()) {
                XSSFCell cell = target.createCell(index);
                cell.setCellValue(iterator.next().getStringCellValue());
            }

            return true;
        } catch (Exception ex) {
            System.out.println(ex.getStackTrace());
        }

        return false;
    }

    /**
     * @param sourceFileKey 是filePath的key
     * @return 读取的行数
     */
    @Override
    public List<XSSFRow> readHeader(String sourceFileKey) {
        List<XSSFRow> header = new ArrayList<>();

        XSSFWorkbook sourceWorkbook = getSourceWorkbook(sourceFileKey);

        XSSFSheet sheet = sourceWorkbook.getSheetAt(0);

        for (int index = 0; index < headerCountMapping.get(sourceFileKey); index++) {
            header.add(sheet.getRow(index));
        }

        return header;
    }

    @Override
    public void writeHeader(String country, String sourceFileKey, List<XSSFRow> rows) {
        List<XSSFRow> dataBody = rows;

        XSSFWorkbook sourceWorkbook = getOutputWorkbook(country, sourceFileKey);

        XSSFSheet sheet = sourceWorkbook.getSheetAt(0);

        int index = sheet.getLastRowNum();

        if (index != 0)
            return;

        for (XSSFRow row : rows) {
            XSSFRow newRow = sheet.createRow(index);
            insertRow(row, newRow);
            index++;
        }
    }

    @Override
    public List<XSSFRow> readDataBody(String sourceFileKey, int pageIndex, int pageCount) {

        List<XSSFRow> dataBody = new ArrayList<>();

        XSSFWorkbook sourceWorkbook = getSourceWorkbook(sourceFileKey);

        XSSFSheet sheet = sourceWorkbook.getSheetAt(0);

        if (pageCount == 0) {
            for (int index = 0; index < pageIndex * pageCount; index++) {
                dataBody.add(sheet.getRow(index));
            }
        } else {
            for (int index = 0; index < sheet.getLastRowNum(); index++) {
                dataBody.add(sheet.getRow(index));
            }
        }

        return dataBody;

    }

    @Override
    public void writeDataBody(String country, String sourceFileKey, List<XSSFRow> rows) {

        List<XSSFRow> dataBody = rows;

        XSSFWorkbook sourceWorkbook = getOutputWorkbook(country, sourceFileKey);

        XSSFSheet sheet = sourceWorkbook.getSheetAt(0);

        int index = sheet.getLastRowNum();

        for (XSSFRow row : rows) {
            XSSFRow newRow = sheet.createRow(index);
            insertRow(row, newRow);
            index++;
        }

    }

    @Override
    public boolean saveAllWorkbook() {
        for (String fullpath : workbook.keySet()) {

            try {
                saveWorkbook(fullpath);
            } catch (Exception e) {
                e.printStackTrace();
                return false;
            }

        }

        return true;
    }

    /**
     * @param country       如国家或者是No group
     * @param sourceFileKey 源文件的文件名，只需要文件名即可
     * @return 保存成功是true 失败是false
     */
    @Override
    public boolean saveSingleWorkbook(String country, String sourceFileKey) {

        String fullpath = getFullPath(country, sourceFileKey);

        try {
            saveWorkbook(fullpath);
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }

        return true;
    }

    @Override
    public void saveWorkbook(String fullpath) throws Exception {
        OutputStream outputStream = null;
        try {

            outputStream = new FileOutputStream(new File(fullpath));
            XSSFWorkbook xssfWorkbook = workbook.get(fullpath);

            xssfWorkbook.write(outputStream);

        } finally {
            if (outputStream != null) {
                try {
                    outputStream.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private String getFullPath(String country, String sourceFileKey) {
        String sourceFileFullPath = sourceFilePropertiesPath.getProperty(sourceFileKey);
        String sourceFileName = sourceFileFullPath.substring(sourceFileFullPath.lastIndexOf(File.separatorChar));
        String outputFileRootPath = configInfo.getProperty("outputFileRootPath");
        return outputFileRootPath + File.separatorChar + country + " " + sourceFileName;
    }
}
