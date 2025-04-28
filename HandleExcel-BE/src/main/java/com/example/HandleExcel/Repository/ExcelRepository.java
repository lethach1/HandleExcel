package com.example.HandleExcel.Repository;

import com.example.HandleExcel.Entity.SalesDataEntity;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

@Repository
public interface ExcelRepository extends JpaRepository<SalesDataEntity, Integer> {
}
