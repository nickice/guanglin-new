package com.guanglin.pptGen.datasource;

import com.guanglin.pptGen.exception.DataSourceException;
import com.guanglin.pptGen.model.Project;

/**
 * Created by pengyao on 01/06/2017.
 */
public abstract class DataSourceBase {

    protected Project project;

    protected DataSourceBase(Project project) {
        this.project = project;
    }

    public abstract Project mapProjectItemData() throws DataSourceException;

}
