package com.example.HandleExcel;

import lombok.Getter;
import lombok.Setter;

import java.util.List;

@Getter
@Setter
public class Errors {
    private List<String> errorMessage;
    private boolean errorFlg;

    public void addErrorMessage(String string) {
        errorMessage.add(string);
    }
}

