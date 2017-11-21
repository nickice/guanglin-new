package com.guanglin.pptGen.model;

import lombok.Data;

import java.util.Queue;

/**
 * PPT Template
 * <p>
 * It is used to save the template slides to create the new ppt file
 * <p>
 * Created by pengyao on 31/05/2017.
 */
public @Data
class PPTTemplate<T> {

    private String templateFileName;

    private Queue<T> slidesQueue;

    private String outputFileName;

    private String outputFilePath;

}
