package com.example.HandleExcel.DTO;

import com.example.HandleExcel.Errors;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class ImportOrderStatus {
    private List<String> errorLine;
    private Errors errors;
    private int status;
    private List<SalesDataDTO> listSalesDataDTO;
}