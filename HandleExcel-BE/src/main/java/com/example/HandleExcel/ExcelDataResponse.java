package com.example.HandleExcel;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.Data;

@Data
public class ExcelDataResponse {
    private String type;
    private int code;
    private String message;
    @JsonInclude(JsonInclude.Include.NON_NULL)
    private Object errors;
    @JsonInclude(JsonInclude.Include.NON_NULL)
    private Object data;

    public ExcelDataResponse() {

    }

    public ExcelDataResponse(String message) {
        this.message = message;
    }
}
