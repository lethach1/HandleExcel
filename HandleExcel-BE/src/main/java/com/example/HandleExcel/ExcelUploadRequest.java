package com.example.HandleExcel;

import org.springframework.web.multipart.MultipartFile;

public class ExcelUploadRequest {

    private MultipartFile file;

    public MultipartFile getFile() {
        return file;
    }

    public void setFile(MultipartFile file) {
        this.file = file;
    }
}

