package com.guanglin.pptGen.datasource;

import com.guanglin.pptGen.datasource.excel.XlsxDataSource;
import com.guanglin.pptGen.exception.DataSourceException;
import com.guanglin.pptGen.model.Project;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.IOException;

import static com.guanglin.pptGen.Constants.DATASOUCE_EXCEL_FILENAME;
import static com.guanglin.pptGen.Constants.DATASOURCE_EXCEL_FOLDER;
import static com.guanglin.pptGen.Constants.PROS;

/**
 * Created by pengyao on 01/06/2017.
 */
public class DataSourceFactory {

    public static Project loadProjectData(String datasourceType, Project project) throws
            IOException,
            InvalidFormatException,
            DataSourceException {

        DataSourceBase datasource = null;

        switch (datasourceType) {
            case "excel":
                String excelFilePath = (String) PROS.get(DATASOURCE_EXCEL_FOLDER) + (String) PROS.get(DATASOUCE_EXCEL_FILENAME);
                FileInputStream fileInputStream = new FileInputStream(excelFilePath);
                datasource = new XlsxDataSource(project, fileInputStream);
                break;
            case "database":
                break;
        }
        if(datasource != null) {
            datasource.mapProjectItemData();
        }

        return project;
    }

}
