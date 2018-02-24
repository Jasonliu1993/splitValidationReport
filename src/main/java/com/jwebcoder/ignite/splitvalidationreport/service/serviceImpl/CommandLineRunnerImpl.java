package com.jwebcoder.ignite.splitvalidationreport.service.serviceImpl;

import com.jwebcoder.ignite.splitvalidationreport.service.ExcelOperator;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.boot.CommandLineRunner;
import org.springframework.stereotype.Service;

import java.io.File;
import java.util.Properties;
import java.util.Set;

@Service
public class CommandLineRunnerImpl implements CommandLineRunner{

    @Autowired
    private ExcelOperator excelOperator;

    @Autowired
    @Qualifier("configInfo")
    private Properties configInfo;

    private String outputFileRootPath;

    @Autowired
    private Properties sourceFilePropertiesPath;

    @Override
    public void run(String... strings) throws Exception {

        Set<String> key = sourceFilePropertiesPath.stringPropertyNames();


    }

}
