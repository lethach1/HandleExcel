package com.example.HandleExcel;

import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;

@Data
@NoArgsConstructor
public class SalesData {
    private String countryCode;
    private String companyCode;
    private String storeCode;
    private String businessDate;
    private BigDecimal salesAmount;
    private BigDecimal pax;
    private BigDecimal salesBudget;
    private BigDecimal paxBudget;
    private String weather;
    private BigDecimal deliverySalesAmount;
    private BigDecimal deliveryPax;
    private BigDecimal eatinSalesAmount;
    private BigDecimal eatinPax;
    private BigDecimal takeoutSalesAmount;
    private BigDecimal takeoutPax;
    private BigDecimal otherSalesAmount;
    private BigDecimal otherPax;

}


