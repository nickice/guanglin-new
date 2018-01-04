package com.guanglin.pptGen.utils;

import java.lang.reflect.Method;

public class ObjectUtils {
    public static void copyPropertyByName(final String propertyName, Object source, Object dest) throws Exception {

        if (source.getClass() != dest.getClass()) {
            throw new Exception("复制属性失败，目标和源对象不匹配");
        }

        Method getMethod = source.getClass().getMethod("get" + propertyName);
        if (getMethod.invoke(source) != null) {
            Method setMethod = dest.getClass().getMethod("set" + propertyName, getMethod.invoke(source).getClass());
            setMethod.invoke(dest, getMethod.invoke(source));

        }
    }
}
