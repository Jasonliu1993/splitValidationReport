package com.jwebcoder.ignite.splitvalidationreport.config;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.Configuration;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.Properties;

@Configuration
public class PropertiesConfig {

    @Value("${customProperties.configPropertiesPath}")
    private String configPropertiesPath;

    @Bean(name = "configInfo")
    public Properties getConfigInfo() {
        Properties configInfo = new Properties();

        try (FileInputStream fileInputStream = new FileInputStream(new File(configPropertiesPath))) {

            configInfo.load(fileInputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return configInfo;
    }

    @Bean(name = "sourceFilePropertiesPath")
    public Properties getSourceFilePropertiesPath(Properties configInfo) {
        Properties sourceFilePropertiesPath = new Properties();
        try (FileInputStream fileInputStream = new FileInputStream(new File(configInfo.getProperty("sourceFilePropertiesPath")))) {

            sourceFilePropertiesPath.load(fileInputStream);

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        return sourceFilePropertiesPath;
    }

    @Bean(name = "countryMapping")
    public Map<String, String> getCountryMapping(Properties configInfo) throws IOException {
        Map<String, String> countryMapping = new HashMap<>();

        FileInputStream inputStream = null;
        XSSFWorkbook mappingWorkbook = null;

        try {
            inputStream = new FileInputStream(new File(configInfo.getProperty("countryMappingPath")));

            mappingWorkbook = new XSSFWorkbook(inputStream);


            XSSFSheet mappingSheet = mappingWorkbook.getSheetAt(0);

            for (Iterator rowIterator = mappingSheet.iterator(); rowIterator.hasNext(); ) {

                XSSFRow row = (XSSFRow) rowIterator.next();

                String key = row.getCell(1).getStringCellValue();

                String value = row.getCell(0).getStringCellValue();

                countryMapping.put(key, value);

            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
        }

        return countryMapping;
    }

    @Bean(name = "headerCountMapping")
    public Map<String, Integer> getHeaderCountMapping(Properties configInfo) throws IOException {
        Map<String, Integer> headerCountMapping = new HashMap<>();

        FileInputStream inputStream = null;
        XSSFWorkbook mappingWorkbook = null;

        try {
            inputStream = new FileInputStream(new File(configInfo.getProperty("headerCountMappingPath")));

            mappingWorkbook = new XSSFWorkbook(inputStream);


            XSSFSheet mappingSheet = mappingWorkbook.getSheetAt(0);

            for (Iterator rowIterator = mappingSheet.iterator(); rowIterator.hasNext(); ) {

                XSSFRow row = (XSSFRow) rowIterator.next();

                String key = row.getCell(0).getStringCellValue();

                Integer value = Integer.valueOf(row.getCell(1).getStringCellValue());

                headerCountMapping.put(key, value);

            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
        }

        return headerCountMapping;
    }


}
