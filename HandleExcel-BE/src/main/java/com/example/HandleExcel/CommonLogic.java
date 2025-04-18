package com.example.HandleExcel;

import lombok.Getter;

@Getter
public class CommonLogic {

    public static final String FORMAT_DATE_YYYYMMDD = "yyyyMMdd";

    public static final String PLAN_RESULT_FLAG = "1";

    public static final String SUFFIX_BUDGET = "Budget";

    public static final String SUFFIX_SALES = "Sales";

    public static final String SUFFIX_TC = "TC";

    public static final String END_ROW = "400";

    public static final String START_COLUMN = "3";

    public static String END_COLUMN = "61";


    public static String trimSheetName(String sheetname){
      if(sheetname == null) return null;
        return sheetname.trim();
    }

//    Function to normalize sheet names by converting full-size characters to half-size
    public static String convertToHalfSize(String sheetName) {
        if (sheetName == null) {
            return null;
        }

        StringBuilder normalized = new StringBuilder();
        for (int i = 0; i < sheetName.length(); i++) {
            char ch = sheetName.charAt(i);

            // Check if the character is full-size letter or number
            if (ch >= '！' && ch <= '～') {
                // Convert full-size characters to half-size
                ch = (char) (ch - 0xFEE0);
            }
            normalized.append(ch);
        }

        return normalized.toString();
    }

    public static boolean validateSheetName(String sheetName, String suffix){
        // Check if sheetName or suffix is empty
        if (sheetName.isEmpty() || suffix.isEmpty()) {
            return false;
        }
        // Regular expression to check year from 2000 to 3000 and valid keyword
        String regex = "^(200\\d|20[1-9]\\d|2[1-9]\\d{2}|3000) ?" + suffix + "$";

        // Check if sheetName does not match the required format
        return sheetName.matches(regex);
    }
}
