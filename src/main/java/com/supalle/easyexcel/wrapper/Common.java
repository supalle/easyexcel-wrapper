package com.supalle.easyexcel.wrapper;

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


}
