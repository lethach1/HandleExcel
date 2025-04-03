package com.example.HandleExcel;

import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

@RestController
@RequestMapping("api/v1")  // Định nghĩa URL gốc
public class MainController {

    @GetMapping("/sales")
    public String getSales() {
        return "Danh sách người dùng";
    }

    @PostMapping("/import")
    public String importMediumExcel() {
        return "Danh sách người dùng";
    }
}



