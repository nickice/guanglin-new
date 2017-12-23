package com.guanglin.pptGen.exception;

/**
 * Created by pengyao on 03/06/2017.
 */
public class ExcelValidationException extends Exception {

    private final static String MESSAGE_FORMAT = "Excel 文件不符合规则：[错误]%s, [更正]%s";

    public ExcelValidationException(String wrongMsg, String rightMsg) {
        super(String.format(MESSAGE_FORMAT, wrongMsg, rightMsg));
    }
}
