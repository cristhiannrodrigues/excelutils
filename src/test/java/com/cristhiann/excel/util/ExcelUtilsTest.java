package com.cristhiann.excel.util;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.junit.jupiter.api.Test;

import java.io.*;

import static com.cristhiann.excel.util.ExcelUtils.generateFileJson;

public class ExcelUtilsTest {

    private InputStream getFileExcel() throws FileNotFoundException {
        ClassLoader classLoader = getClass().getClassLoader();
        File fileExcel = new File(classLoader.getResource("teste.xlsx").getFile());
        InputStream input = new FileInputStream(fileExcel);
        return input;
    }


    @Test
    void testExcel() throws IOException, InvalidFormatException {
        printJSON(generateFileJson(getFileExcel()));
    }

    private void printJSON(String json) {
        System.out.println(json);
    }

}
