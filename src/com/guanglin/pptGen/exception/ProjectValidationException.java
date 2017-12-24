package com.guanglin.pptGen.exception;

import com.guanglin.pptGen.model.Item;
import lombok.Getter;

import java.util.Map;

public class ProjectValidationException extends Exception {

    @Getter
    private Map<Item, String> invalidItems;

    public ProjectValidationException(final String msg) {
        super(msg);
    }

    public ProjectValidationException(final String msg, Map<Item, String> invalidItems) {
        super(msg);

        this.invalidItems = invalidItems;
    }


}
