package com.guanglin.pptGen.exception;

/**
 * Created by pengyao on 12/06/2017.
 */
public class PPTException extends Exception {

    public PPTException(final String msg) {
        super(msg);
    }

    public PPTException(final String msg, Exception ex) {
        super(msg, ex);
    }
}
