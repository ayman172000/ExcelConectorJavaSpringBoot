package com.example.demoexcel.service;



import com.example.demoexcel.dto.FileDTO;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.*;

@Service
public class Excelservice implements  IExcelservice {

 /*   @Override
    public List<List<Map<String, Object>>> getExcelData(MultipartFile file) throws IOException {
        List<List<Map<String, Object>>> result = new ArrayList<>();
        try {
            Workbook workbook = new XSSFWorkbook(file.getInputStream());
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                Row headerRow = sheet.getRow(0);
                List<String> headers = new ArrayList<>();
                for (Cell cell : headerRow) {
                    headers.add(cell.getStringCellValue());
                }
                List<Map<String, Object>> sheetData = new ArrayList<>();
                for (Row row : sheet) {
                    Map<String, Object> rowData = new LinkedHashMap<>();
                    for (Cell cell : row) {
                        for (int k = 0; k < headers.size(); k++) {
                            //String columnName = getColumnName(cell.getColumnIndex());
                            String columnName = cell.getStringCellValue();
                            switch (cell.getCellType()) {
                                case NUMERIC:
                                    rowData.put(headers.get(k), cell.getNumericCellValue());
                                    break;
                                case BOOLEAN:
                                    rowData.put(headers.get(k), cell.getBooleanCellValue());
                                    break;
                                case STRING:
                                    rowData.put(headers.get(k), cell.getStringCellValue());
                                    break;
                                default:
                                    rowData.put(headers.get(k), null);
                                    break;
                            }
                        }
                    }
                    sheetData.add(rowData);
                }
                result.add(sheetData);
            }
            return result;
        }
        catch (Exception e)
        {
            System.out.println(e);
            return null;
        }
    }


    @Override
    public String getColumnName(int columnNumber) {
        StringBuilder columnName = new StringBuilder();
        while (columnNumber >= 0) {
            int remainder = columnNumber % 26;
            columnName.insert(0, (char) ('A' + remainder));
            columnNumber = (columnNumber / 26) - 1;
        }
        return columnName.toString();
    }*/





    /*public FileDTO readExcelFile(MultipartFile file) throws IOException {
        FileDTO fileDTO=new FileDTO();
        //FileInputStream inputStream = new FileInputStream(new File(fileName));
        Workbook wb = new XSSFWorkbook(file.getInputStream());
        for(int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            Row headerRow = sheet.getRow(0);
            List<String> tempHeaders = new ArrayList<>();
            for (Cell cell : headerRow) {
                tempHeaders.add(cell.getStringCellValue());
            }
            List<Map<String, Object>> data = new ArrayList<>();
            String sheetName = sheet.getSheetName();
            for (Row row : sheet) {
                Map<String, Object> rowData = new HashMap<>();
                for(Cell cell : row) {
                    String value="";
                    if(cell.getCellType()==CellType.NUMERIC)
                        value=String.valueOf(cell.getNumericCellValue());
                    else if(cell.getCellType()==CellType.BOOLEAN)
                        value=String.valueOf(cell.getBooleanCellValue());
                    else if(cell.getCellType()==CellType.STRING)
                        value=cell.getStringCellValue();
                    rowData.put(cell.getColumnIndex() + "", value);
                }
                data.add(rowData);
            }
            excelData.add(new HashMap<String, Object>() {{
                put("sheetName", sheetName);
                put("data", data);
            }});
            List<String> headersList = new ArrayList<>();
            Row firstRow = sheet.getRow(0);
            for(Cell cell : firstRow) {
                headersList.add(cell.getStringCellValue());
            }
            String[] headers = headersList.toArray(new String[headersList.size()]);
            this.headers.put(sheetName, headers);
            this.feuilles.add(sheetName);
        }
        fileDTO.setExcelData(this.excelData);
        fileDTO.setHeaders(this.headers);
        fileDTO.setFeuilles(this.feuilles);
        System.out.println(fileDTO);
        return fileDTO;
    }*/

    List<Map<String, Object>> excelData = new ArrayList<>();
    Map<String, String[]> headers = new HashMap<>();
    List<String> feuilles = new ArrayList<>();
    @Override
    public FileDTO readExcelFile(MultipartFile file) throws IOException{
        FileDTO fileDTO=new FileDTO();
        //FileInputStream inputStream = new FileInputStream(new File(fileName));
        Workbook wb = new XSSFWorkbook(file.getInputStream());
        for(int i = 0; i < wb.getNumberOfSheets(); i++) {
            Sheet sheet = wb.getSheetAt(i);
            List<Map<String, Object>> data = new ArrayList<>();
            String sheetName = sheet.getSheetName();
            Row firstRow = sheet.getRow(0);
            String[] headers = new String[firstRow.getLastCellNum()];
            for(Cell cell : firstRow) {
                headers[cell.getColumnIndex()] = cell.getStringCellValue();
            }
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                Map<String, Object> rowData = new LinkedHashMap<>();
                for(Cell cell : row) {
                    int colIndex = cell.getColumnIndex();
                    String header = headers[colIndex];
                    String value="";
                    if(cell.getCellType()== CellType.STRING)
                        value=cell.getStringCellValue();
                    else if(cell.getCellType()== CellType.NUMERIC)
                        value=String.valueOf(cell.getNumericCellValue());
                    else if(cell.getCellType()== CellType.BOOLEAN)
                        value=String.valueOf(cell.getBooleanCellValue());
                    rowData.put(header, value);
                }
                data.add(rowData);
            }
            excelData.add(new HashMap<String, Object>() {{
                put("sheetName", sheetName);
                put("data", data);
            }});
            this.headers.put(sheetName, headers);
            this.feuilles.add(sheetName);
        }
        fileDTO.setFeuilles(this.feuilles);
        fileDTO.setExcelData(this.excelData);
        fileDTO.setHeaders(this.headers);
        return fileDTO;
    }
}

