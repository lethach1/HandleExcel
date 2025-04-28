package com.example.HandleExcel.DTO;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class SalesDataDTO {

    private String storeCode;

    private String businessDate;

    private String salesAmount;

    private String pax;

    private String salesBudget;

    private String paxBudget;

}
