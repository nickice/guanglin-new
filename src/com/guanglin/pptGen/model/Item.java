package com.guanglin.pptGen.model;

import lombok.Data;

import java.util.Map;
import java.util.Stack;

/**
 * Created by pengyao on 30/05/2017.
 */
public @Data
class Item {
    private Capture outsideCapture;
    private String description;
    private Stack<Capture> insideCaptures;
    private Map<String, String> fields;

    public int getInsideImageCounts() {
        if (this.insideCaptures == null) {
            return 0;
        }

        return this.insideCaptures.size();
    }

}
