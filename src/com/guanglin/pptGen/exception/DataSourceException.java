package com.guanglin.pptGen.exception;

/**
 * Created by pengyao on 12/06/2017.
 */
public class DataSourceException extends Exception {

    public DataSourceException(final String msg) {
        super(msg);
    }

    public DataSourceException(Exception e) {
        super(e);
    }

    public DataSourceException(String msg, String cellPosition) {
        this(String.format("%s\n 请检查Excel文件内容，单元格:%s", msg, cellPosition ));

    }
}
