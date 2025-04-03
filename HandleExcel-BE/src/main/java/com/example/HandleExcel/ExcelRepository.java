package com.example.HandleExcel;

import org.springframework.data.jpa.repository.JpaRepository;

public interface ExcelRepository extends JpaRepository<SalesData, Integer> {
}
