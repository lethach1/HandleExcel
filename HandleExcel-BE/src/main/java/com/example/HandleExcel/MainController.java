package com.example.HandleExcel;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import java.io.IOException;

@RestController
@RequestMapping("api/v1")  // Định nghĩa URL gốc
public class MainController {

    @Autowired
    private ExcelService excelService;

    @GetMapping("/sales")
    public String getSales() {
        return "Danh sách người dùng";
    }

    @PostMapping("/upload")
    public ResponseEntity<String> uploadSmallExcel(@ModelAttribute ExcelUploadRequest request) {

        ImportOrderStatus dto;
        try {
            dto = excelService.uploadSmallExcel(request);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        ExcelDataResponse excelDataResponse = new ExcelDataResponse();
        excelDataResponse.setType("success");
        excelDataResponse.setMessage("インポートは処理中です。");
        excelDataResponse.setCode(200);
        excelDataResponse.setData(dto);

        return ResponseEntity.ok("File uploaded successfully");
    }
}



