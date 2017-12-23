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

        LOGGER.info("开始读取数据内容.");
        switch (datasourceType) {
            case "excel":
                String excelFilePath = project.getProjectPath() + PRO_DATA_PATH;
                LOGGER.info("数据文件地址:" + excelFilePath);

                File excelFile = new File(excelFilePath);
                File[] excelFiles = excelFile.listFiles();
                if (excelFile == null || excelFiles.length == 0) {
                    throw new ProjectException("没有在项目目录下找到数据文件。");
                }

                FileInputStream fileInputStream = new FileInputStream(excelFiles[0]);
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
