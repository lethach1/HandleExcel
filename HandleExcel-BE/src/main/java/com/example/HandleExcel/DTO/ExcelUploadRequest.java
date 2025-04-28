package com.example.HandleExcel.DTO;

import lombok.Getter;
import lombok.Setter;
import org.springframework.web.multipart.MultipartFile;

@Getter
@Setter
public class ExcelUploadRequest {

    private MultipartFile file;
    private String planFlag;
}



