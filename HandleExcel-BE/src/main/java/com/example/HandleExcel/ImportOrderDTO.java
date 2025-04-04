package com.example.HandleExcel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class ImportOrderDTO {
    private int numberOfProcess;
    private int numberOfUpdate;
    private int numberOfInsert;
    private int numberOfError;
    private int numberOfWarning;
    private List<String> errorLine;
    private List<String> warningLine;
    private int status;
    private Long idQueue;
}