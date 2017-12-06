package com.guanglin.pptGen.model;

import com.guanglin.pptGen.model.pptConfig.PPTConfig;
import lombok.Data;

import java.util.List;

/**
 * Project
 * <p>
 * Each project will have multiple items
 * Created by pengyao on 30/05/2017.
 */
public @Data
class Project {
    /**
     * 1. one project has one owner
     * 2. one project has multiple items
     * 3. one project has one folder path
     */

    private String owner;
    private List<Item> items;
    private String captureFolderPath;
    private PPTConfig pptConfig;
}
