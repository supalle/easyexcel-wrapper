package com.supalle.easyexcel.wrapper;

import com.alibaba.excel.enums.CellDataTypeEnum;
import com.alibaba.excel.metadata.CellData;

import java.util.Collection;

public class Common {

    public static boolean isNull(Object object) {
        return object == null;
    }

    public static boolean isEmpty(String s) {
        return isNull(s) || s.trim().isEmpty();
    }

    public static boolean isNotEmpty(String s) {
        return !isEmpty(s);
    }

    public static boolean isEmpty(Collection<?> collection) {
        return isNull(collection) || collection.isEmpty();
    }

    public static boolean isNotEmpty(Collection<?> collection) {
        return !isEmpty(collection);
    }

    public static boolean isNull(CellData cellData) {
        if (cellData == null) {
            return true;
        }
        CellDataTypeEnum type = cellData.getType();
        if (type == null) {
            return true;
        }
        switch (type) {
            case STRING:
            case ERROR:
                return cellData.getStringValue() == null;
            case NUMBER:
                return cellData.getNumberValue() == null;
            case BOOLEAN:
                return cellData.getBooleanValue() == null;
            default:
        }
        return false;
    }

    public static boolean isEmpty(CellData cellData) {
        if (cellData == null) {
            return true;
        }
        CellDataTypeEnum type = cellData.getType();
        if (type == null || type == CellDataTypeEnum.EMPTY) {
            return true;
        }
        switch (type) {
            case STRING:
            case ERROR:
                return isEmpty(cellData.getStringValue());
            case NUMBER:
                return cellData.getNumberValue() == null;
            case BOOLEAN:
                return cellData.getBooleanValue() == null;
            default:
        }
        return false;
    }

    public static boolean isNotEmpty(CellData cellData) {
        return !isEmpty(cellData);
    }

}
