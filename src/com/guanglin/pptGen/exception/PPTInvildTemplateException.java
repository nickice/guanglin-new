package com.guanglin.pptGen.exception;

/**
 * PPT Invalid Template Exception
 * <p>
 * Created by pengyao on 31/05/2017.
 */
public class PPTInvildTemplateException extends PPTException {

    public PPTInvildTemplateException(final String msg) {
        super(msg);
    }

    public PPTInvildTemplateException(final String msg, Exception ex) {
        super(msg, ex);
    }

}
