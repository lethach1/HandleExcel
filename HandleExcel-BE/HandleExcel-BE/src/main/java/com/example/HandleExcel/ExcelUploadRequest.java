package com.example.HandleExcel;

import lombok.Data;
import org.springframework.web.multipart.MultipartFile;

@Data
public class ExcelUploadRequest {
    private MultipartFile file;
    private String planFlag;
}
