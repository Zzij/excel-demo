package com.zz.excel.util;

import java.text.SimpleDateFormat;
import java.util.Date;

public class DateFormatUtil {


    static final String format1 = "yyyy-MM-dd HH:mm:ss";

    static SimpleDateFormat sdf = new SimpleDateFormat(format1);

    public static String getStrDate(Date date){
        return sdf.format(date);
    }
}
