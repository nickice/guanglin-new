package com.guanglin.pptGen.exception;

import lombok.NonNull;

public class ProjectValidationException extends Exception {

    public ProjectValidationException(@NonNull String msg) {
        super(msg);
    }
}
