package com.example.demoexcel.dto;

import lombok.Data;

@Data
public class ExcelFileDTO {
    private String fileChecksum;
    private Long fileSize;
    private String fileName;
    private String jsonValue;
}
