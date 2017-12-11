package com.guanglin.pptGen.model;

import com.guanglin.pptGen.Constants;
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

    private String name;
    private String projectPath;
    private String projectPPTTemplatePath;
    private String projectOutputPath;
    private List<Item> items;
    private String captureFolderPath;
    private PPTConfig pptConfig;

    public String getprojectPPTTemplatePath() {
        return this.projectPath + Constants.PRO_TEMPLATE_PATH;
    }

    public String getProjectOutputPath() {
        return this.projectPath + Constants.PRO_OUTPUT_PATH;
    }
}
