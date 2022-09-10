package com.cristhiann.excel.util;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.commons.compress.archivers.dump.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

public class ExcelUtils {

    private static Gson gson = new GsonBuilder().create();

    private static Iterator<Row> openExcel(InputStream input) throws IOException, InvalidFormatException {
        try( XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(input)) ) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            return sheet.iterator();
        } catch (org.apache.poi.openxml4j.exceptions.InvalidFormatException e) {
            throw new RuntimeException(e);
        }
    }


}
