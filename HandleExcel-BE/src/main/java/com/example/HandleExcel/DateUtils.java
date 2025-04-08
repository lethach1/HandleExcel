package com.example.HandleExcel;

import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

public class DateUtils {
    public static String formateDate(Date date, String format) {
        return date.toInstant()
                    .atZone(ZoneId.systemDefault())
                    .toLocalDateTime()
                    .format(DateTimeFormatter.ofPattern(format));
    }
}
