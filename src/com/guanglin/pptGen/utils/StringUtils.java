package com.guanglin.pptGen.utils;

import com.google.common.base.Strings;

public class StringUtils {

    public static String trim(String s) {
        if (Strings.isNullOrEmpty(s)) {
            return s;
        }

        return s.trim().replace(" ", "").replace("\n", "").replace("\r", "");
    }
}
