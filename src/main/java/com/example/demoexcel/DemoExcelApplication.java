package com.example.demoexcel;


import org.apache.poi.ss.usermodel.*;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@SpringBootApplication
public class DemoExcelApplication {

    public static void main(String[] args) {
        /*FichierExcel fichierExcel=new FichierExcel("uploads/Classeur1.xlsx");
        System.out.println("fichier:"+fichierExcel.getFeuilles());*/
        SpringApplication.run(DemoExcelApplication.class, args);
    }
}
