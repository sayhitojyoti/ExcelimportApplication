package com.Practice.ExcelImportApplication.controller;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import com.Practice.ExcelImportApplication.service.ExcelService;

@RestController
@RequestMapping("/api/excel")
public class ExcelController {

    @Autowired
    private ExcelService excelService;

    @CrossOrigin(origins = "*")
    @PostMapping("/upload")
    public ResponseEntity<?> uploadExcel(@RequestParam("file") MultipartFile[] files) {
        for (MultipartFile file : files) {
            excelService.processExcelFile(file); 
        }
        return ResponseEntity.ok(" Files are being processed in the background.");
    }

}
