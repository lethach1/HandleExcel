package com.example.HandleExcel.Controller;

import com.example.HandleExcel.DTO.ExcelDataResponse;
import com.example.HandleExcel.Service.ExcelService;
import com.example.HandleExcel.DTO.ExcelUploadRequest;
import com.example.HandleExcel.DTO.ImportOrderStatus;
import jakarta.servlet.http.HttpSession;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
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
    public ResponseEntity<ExcelDataResponse> uploadSmallExcel(@ModelAttribute ExcelUploadRequest request, HttpSession session) {

        ImportOrderStatus dto;
        try {
            dto = excelService.uploadSmallExcel(request);
            session.setAttribute("lisSalesDataDTO", dto.getListSalesDataDTO());

            ExcelDataResponse excelDataResponse = new ExcelDataResponse();
            excelDataResponse.setType("success");
            excelDataResponse.setMessage("The import is in progress.");
            excelDataResponse.setCode(200);
            excelDataResponse.setData(dto);

            return ResponseEntity.ok(excelDataResponse);
        } catch (IOException e) {
            ExcelDataResponse excelDataResponse = new ExcelDataResponse();
            excelDataResponse.setType("error");
            excelDataResponse.setMessage("Error processing file.");
            excelDataResponse.setCode(200);
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body(excelDataResponse);
        }
    }
}



