package com.example.HandleExcel;

import jakarta.persistence.Column;
import jakarta.persistence.Entity;
import jakarta.persistence.Id;
import jakarta.persistence.Table;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.math.BigDecimal;

@Entity
@Data
@NoArgsConstructor
@Table(name = "sales_data")
public class SalesDataEntity {

    @Id
    @Column(name = "country_code", length = 4, nullable = false)
    private String countryCode;

    @Column(name = "company_code", length = 12, nullable = false)
    private String companyCode;

    @Column(name = "store_code", length = 12, nullable = false)
    private String storeCode;

    @Column(name = "business_date", length = 8, nullable = false)
    private String businessDate;

    @Column(name = "sales_amount", precision = 14, scale = 2)
    private BigDecimal salesAmount;

    @Column(name = "pax", precision = 14)
    private BigDecimal pax;

    @Column(name = "sales_budget", precision = 14, scale = 2)
    private BigDecimal salesBudget;

    @Column(name = "pax_budget", precision = 14)
    private BigDecimal paxBudget;

    @Column(name = "weather", length = 20)
    private String weather;

    @Column(name = "delivery_sales_amount", precision = 14, scale = 2)
    private BigDecimal deliverySalesAmount;

    @Column(name = "delivery_pax", precision = 14)
    private BigDecimal deliveryPax;

    @Column(name = "eatin_sales_amount", precision = 14, scale = 2)
    private BigDecimal eatinSalesAmount;

    @Column(name = "eatin_pax", precision = 14)
    private BigDecimal eatinPax;

    @Column(name = "takeout_sales_amount", precision = 14, scale = 2)
    private BigDecimal takeoutSalesAmount;

    @Column(name = "takeout_pax", precision = 14)
    private BigDecimal takeoutPax;

    @Column(name = "other_sales_amount", precision = 14, scale = 2)
    private BigDecimal otherSalesAmount;

    @Column(name = "other_pax", precision = 14)
    private BigDecimal otherPax;
}


