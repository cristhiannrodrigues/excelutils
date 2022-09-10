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

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class ExcelUtils {

    private static Gson gson = new GsonBuilder().create();

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
