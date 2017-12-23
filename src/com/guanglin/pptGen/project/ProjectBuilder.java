package com.guanglin.pptGen.project;

import com.google.common.base.Strings;
import com.guanglin.pptGen.datasource.DataSourceFactory;
import com.guanglin.pptGen.exception.DataSourceException;
import com.guanglin.pptGen.exception.ProjectException;
import com.guanglin.pptGen.exception.ProjectValidationException;
import com.guanglin.pptGen.model.Capture;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.Project;
import lombok.NonNull;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.IOException;

import static com.guanglin.pptGen.Constants.APPCONFIG_PRO_NAME;
import static com.guanglin.pptGen.Constants.DATASOURCE_TYPE;
import static com.guanglin.pptGen.Constants.PROS;
import static com.guanglin.pptGen.Constants.PRO_CONFIG_PATH;
import static com.guanglin.pptGen.Constants.PRO_DATA_PATH;
import static com.guanglin.pptGen.Constants.PRO_IMG_OUTDORR_PATH;
import static com.guanglin.pptGen.Constants.PRO_IMG_PATH;
import static com.guanglin.pptGen.Constants.PRO_OUTPUT_PATH;
import static com.guanglin.pptGen.Constants.PRO_TEMPLATE_PATH;

/**
 * Created by pengyao on 02/07/2017.
 */
public class ProjectBuilder {

    private static Logger LOGGER = LogManager.getLogger("ProjectBuilder");

    public static Project build(String projectPath) throws ProjectValidationException, InvalidFormatException, DataSourceException, ProjectException, IOException {

        Project project = new Project();

        project.setProjectPath(projectPath);
        LOGGER.info("项目目录地址：" + projectPath);

        // assign the project name
        project.setName(PROS.getProperty(APPCONFIG_PRO_NAME));
        LOGGER.info("项目名称：" + PROS.getProperty(APPCONFIG_PRO_NAME));

        // project path
        validateProjectConstuctor(projectPath);

        // assign data source setting
        final String datasourceType = PROS.getProperty(DATASOURCE_TYPE);
        LOGGER.info("项目数据类型：" + datasourceType);

        LOGGER.debug("Parse template config successfully.");

        project = DataSourceFactory.loadProjectData(datasourceType, project);
        loadCaptures(projectPath, project);

        return project;

    }

    private static void validateProjectConstuctor(@NonNull String projectPath) throws ProjectValidationException {

        // 1. the project path must exist
        File projectFolder = new File(projectPath);
        if (!projectFolder.exists() || !projectFolder.isDirectory()) {
            throw new ProjectValidationException("Project Folder doesn't exist");
        }

        // 2. the data folder must exist
        File dataFolder = new File(projectPath + PRO_DATA_PATH);
        if (!dataFolder.exists() || !dataFolder.isDirectory()) {
            throw new ProjectValidationException("Project/data Folder doesn't exist");
        }

        // 3. the images must exist
        File imageFolder = new File(projectPath + PRO_IMG_PATH);
        if (!imageFolder.exists() || !imageFolder.isDirectory()) {
            throw new ProjectValidationException("Project/image Folder doesn't exist");
        }

        // 4. the template file must exist
        File templateFile = new File(projectPath + PRO_TEMPLATE_PATH);
        if (!templateFile.exists()) {
            throw new ProjectValidationException("Project/template.ppt doesn't exist");
        }

        // 5. create the output folder if it doesn't exist
        File outputFolder = new File(projectPath + PRO_OUTPUT_PATH);
        if (!outputFolder.exists()) {
            outputFolder.mkdir();
        }

        // 6. the template file must exist
        File configFile = new File(projectPath + PRO_CONFIG_PATH);
        if (!configFile.exists()) {
            throw new ProjectValidationException("Project/project.ppt doesn't exist");
        }

    }


    private static void loadCaptures(@NonNull String projectPath, @NonNull Project project) throws ProjectValidationException {

        final String projectImageDirPath = projectPath + PRO_IMG_PATH;
        final String itemImageOurdoorPath = projectImageDirPath + "/" + PRO_IMG_OUTDORR_PATH;

        // ensure the item image folder is existing
        File prjectImageDir = new File(projectImageDirPath);
        if (!prjectImageDir.exists()
                || !prjectImageDir.isDirectory()
                || prjectImageDir.listFiles().length == 0) {
            throw new ProjectValidationException(String.format("% Folder doesn't exist", projectImageDirPath));
        }

        for (Item item : project.getItems()) {

            if (Strings.isNullOrEmpty(item.getDescription())) {
                continue;
            }

            // item folder
            String itemImageDirPath = projectImageDirPath + "/" + item.getDescription();
            File itemImageDir = new File(itemImageDirPath);

            // check item image folders
            if (!itemImageDir.exists()
                    || !itemImageDir.isDirectory()
                    || itemImageDir.listFiles().length == 0) {
                throw new ProjectValidationException(String.format("%s image folder is no exist or empty. ", itemImageDirPath));
            }

            for (File f : itemImageDir.listFiles()) {
                // set outdoor image
                if (f.isDirectory()) {
                    if (f.listFiles().length != 1) {
                        throw new ProjectValidationException(String.format("%s image folder has problem", f.getPath()));
                    }

                    File outdoorImage = f.listFiles()[0];
                    item.setOutsideCapture(new Capture(outdoorImage.getName(), outdoorImage.getAbsolutePath()));
                    continue;
                }

                // loop to set indoor image
                item.getInsideCaptures().add(new Capture(f.getName(), f.getAbsolutePath()));
            }
        }

    }

}
