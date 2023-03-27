package com.example.demoexcel.web;



import com.example.demoexcel.dto.FileDTO;
import com.example.demoexcel.service.IExcelservice;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;



@RestController
@CrossOrigin("*")
public class ExcelDataController {
    IExcelservice iExcelservice;

    public ExcelDataController(IExcelservice iExcelservice) {
        this.iExcelservice = iExcelservice;
    }
    @PostMapping("/calculChecksum")
    public FileDTO calculerChecksum(@RequestParam("file") MultipartFile file) throws Exception
    {
        return this.iExcelservice.readExcelFile(file);
    }
}

