package com.cristhiann.excel.util;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.List;

public class ExcelUtils {

    private static Gson gson = new GsonBuilder().create();

    public static String generateFileJson(InputStream input) throws IOException, InvalidFormatException {
        List<LinkedHashMap<String,String>> listLinkedMap = new ArrayList<>();
        List<List<String>> arquivo = retornaDadosArquivo(input);
        List<String> titulo = arquivo.get(0);
        arquivo.remove(0);
        for(List<String> linha : arquivo) {
            LinkedHashMap<String,String> linkedMap = new LinkedHashMap<>();
            for(Integer iterator = 0; iterator < linha.size(); iterator++) {
                linkedMap.put(titulo.get(iterator), linha.get(iterator));
            }
            listLinkedMap.add(linkedMap);
        }
        return gson.toJson(listLinkedMap);
    }

    private static  List<List<String>> retornaDadosArquivo(InputStream input) throws IOException, InvalidFormatException {
        List<List<String>> listCells = new ArrayList<>();
        Iterator<Row> rowIterator = openExcel(input);
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            List<String> cells = new ArrayList<>();
            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                if(cell.getCellType() == CellType.STRING) {
                    cells.add(cell.getStringCellValue());
                } else if(cell.getCellType() == CellType.NUMERIC) {
                    cells.add(String.valueOf(cell.getNumericCellValue()));
                }
            }
            listCells.add(cells);
        }
        return listCells;
    }

    private static Iterator<Row> openExcel(InputStream input) throws IOException, InvalidFormatException {
        try( XSSFWorkbook workbook = new XSSFWorkbook(OPCPackage.open(input)) ) {
            XSSFSheet sheet = workbook.getSheetAt(0);
            return sheet.iterator();
        }
    }


}
