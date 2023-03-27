package com.example.demoexcel.service;

import com.example.demoexcel.dto.FileDTO;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface IExcelservice {
    //List<List<Map<String, Object>>> getExcelData(MultipartFile file) throws IOException;
    //String getColumnName(int columnNumber);

    FileDTO readExcelFile(MultipartFile file) throws IOException;
}
