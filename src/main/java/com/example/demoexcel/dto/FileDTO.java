package com.example.demoexcel.dto;

import lombok.Data;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

@Data
public class FileDTO {
    List<Map<String, Object>> excelData = new ArrayList<>();
    Map<String, String[]> headers = new HashMap<>();
    List<String> feuilles = new ArrayList<>();
}
