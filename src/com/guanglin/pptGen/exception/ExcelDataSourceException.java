package com.guanglin.pptGen.exception;

public class ExcelDataSourceException extends DataSourceException {
    private static final String ERR_MSG = "Excel 数据内容、格式错误：";
    private static final String ERR_1 = ERR_MSG + "1. 不可识别的数据格式，请检查数据内容，单元格：%s";

    public ExcelDataSourceException(String cellPosition) {
        super(String.format(ERR_1, cellPosition));
    }
}
