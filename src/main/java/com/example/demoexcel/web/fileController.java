package com.example.demoexcel.web;

import com.example.demoexcel.dto.ExcelFileDTO;
import com.fasterxml.jackson.core.type.TypeReference;
import com.fasterxml.jackson.databind.ObjectMapper;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.xml.bind.DatatypeConverter;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.security.MessageDigest;
import java.util.*;

@RestController
@CrossOrigin("*")
public class fileController {

    //private String checksumTest="D8B4E8CE3792DB94B5E30EED6E0C4E9516C3098960B4AAF2908603C5C2E4AAEC";
    @PostMapping("/calculChecksum")
    public ExcelFileDTO calculerChecksum(@RequestParam("file") MultipartFile file) throws Exception {
            MessageDigest md = MessageDigest.getInstance("SHA-256");
            byte[] checksum = md.digest(file.getBytes());
            String checksumHex = DatatypeConverter.printHexBinary(checksum);
            //if(checksumHex.equals(checksumTest))
                //throw new Exception("fichier existe deja");
            byte[] bytes = file.getBytes();
            // Définir le chemin de destination du fichier
            Path path = Paths.get("uploads/" + file.getOriginalFilename());
            // Écrire le contenu du fichier dans le fichier de destination
            Files.write(path, bytes);

            ExcelFileDTO excelFileDTO=new ExcelFileDTO();
            excelFileDTO.setFileChecksum(checksumHex);
            excelFileDTO.setFileName(file.getOriginalFilename());
            excelFileDTO.setFileSize(file.getSize());
        String json = this.convertExcelToJson(file);
        System.out.println("json value:"+json);
        excelFileDTO.setJsonValue(json);
        return excelFileDTO;
    }

    public String convertExcelToJson(MultipartFile file) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream()); // ou new HSSFWorkbook(file.getInputStream()) pour un fichier .xls
        List<Map<String, String>> allRows = new ArrayList<>();
// boucler sur toutes les feuilles du classeur
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            System.out.println("iteration N°:"+i);
            Sheet sheet = workbook.getSheetAt(i);
            List<Map<String, String>> rows = new ArrayList<>();

            // récupérer la première ligne pour les en-têtes de colonnes
            Row headerRow = sheet.getRow(0);
            List<String> headers = new ArrayList<>();
            for (Cell cell : headerRow) {
                headers.add(cell.getStringCellValue());
            }

            // boucler sur chaque ligne de la feuille et récupérer les données de chaque cellule
            for (int j = 1; j <= sheet.getLastRowNum(); j++) {
                Row row = sheet.getRow(j);
                Map<String, String> rowData = new HashMap<>();
                for (int k = 0; k < headers.size(); k++) {
                    Cell cell = row.getCell(k);
                    String header = headers.get(k);
                    String value = "";
                    if (cell != null) {
                        value = cell.getStringCellValue();
                    }
                    rowData.put(header, value);
                }
                rows.add(rowData);
            }
            allRows.addAll(rows);
        }
        ObjectMapper mapper = new ObjectMapper();
        String json = mapper.writeValueAsString(allRows);
        //ExcelFileDTO excelFileDTO = mapper.readValue(json, ExcelFileDTO.class);
        //List<ExcelFileDTO> excelFileDTOS = mapper.readValue(json, new TypeReference<List<ExcelFileDTO>>(){});
        //System.out.println("list object:====>"+excelFileDTOS);
        return json;

    }


}
