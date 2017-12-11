package com.guanglin.pptGen.datasource;

import com.guanglin.pptGen.datasource.excel.XlsxDataSource;
import com.guanglin.pptGen.exception.DataSourceException;
import com.guanglin.pptGen.exception.ProjectException;
import com.guanglin.pptGen.model.Project;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;

import static com.guanglin.pptGen.Constants.DATASOURCE_EXCEL_FOLDER;
import static com.guanglin.pptGen.Constants.PROS;
import static com.guanglin.pptGen.Constants.PRO_DATA_PATH;

/**
 * Created by pengyao on 01/06/2017.
 */
public class DataSourceFactory {

    private static Logger LOGGER = LogManager.getLogger("DataSourceFactory");

    public static Project loadProjectData(String datasourceType, Project project) throws
            IOException,
            InvalidFormatException,
            DataSourceException, ProjectException {

        DataSourceBase datasource = null;

        switch (datasourceType) {
            case "excel":
                File excelFilePath = new File(project.getProjectPath() + PRO_DATA_PATH);
                File[] excelFile = excelFilePath.listFiles();
                if (excelFile == null || excelFile.length == 0) {
                    throw new ProjectException("no excel file found");
                }

                FileInputStream fileInputStream = new FileInputStream(excelFile[0]);
                datasource = new XlsxDataSource(project, fileInputStream);
                break;
            case "database":
                break;
        }
        if (datasource != null) {
            datasource.mapProjectItemData();
        }

        return project;
    }

}
