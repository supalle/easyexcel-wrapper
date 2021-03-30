package com.supalle.easyexcel.wrapper;

/**
 * Excel 操作异常
 * 2019/6/15
 *
 * @author WeiBQ
 */
public class ExcelException extends RuntimeException {

    public ExcelException(String message) {
        super(message);
    }

    public ExcelException(String message, Throwable cause) {
        super(message, cause);
    }
}
