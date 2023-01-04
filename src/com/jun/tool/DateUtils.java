package com.jun.tool;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

/**
 * 时间工具类
 * 
 * @author ruoyi
 */
public class DateUtils{


    public static DateTimeFormatter date = DateTimeFormatter.ofPattern("yyyy年MM月dd日");
    public static DateTimeFormatter dateNo = DateTimeFormatter.ofPattern("yyMMdd");
    public static DateTimeFormatter time = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    public static String toString(LocalDateTime time, DateTimeFormatter data){
        return data.format(time);
    }

    public static LocalDateTime toLocalDateTime(String time,DateTimeFormatter data){
        return LocalDateTime.parse(time,data);
    }

}
