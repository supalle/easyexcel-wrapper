package com.supalle.easyexcel.wrapper;

import java.time.format.DateTimeFormatter;

public class ExcelConstant {

    public static final DateTimeFormatter DATE_TIME = DateTimeFormatter.ofPattern("yyyy/M/d HH:mm:ss");

    public static final DateTimeFormatter DATE = DateTimeFormatter.ofPattern("yyyy/M/d");

    public static final DateTimeFormatter TIME = DateTimeFormatter.ofPattern("HH:mm:ss");

}
