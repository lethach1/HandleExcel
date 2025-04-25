package com.example.HandleExcel;

import io.micrometer.common.util.StringUtils;

public class CheckUtil {

    /**Need to custom isEmpty method because of full-width space in japanese **/
    public static boolean isEmpty(String str) {
        if (StringUtils.isBlank(str)) {
            return true;
        }
        String val = (String) str;
        val = val.replaceAll("　", "");
        if (val.trim().length() == 0) {
            return true;
        }
        return false;
    }
}
