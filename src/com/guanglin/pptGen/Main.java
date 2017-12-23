package com.guanglin.pptGen;

import com.google.common.base.Strings;
import com.guanglin.pptGen.datasource.DataSourceFactory;
import com.guanglin.pptGen.exception.ProjectValidationException;
import com.guanglin.pptGen.model.Capture;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.Project;
import com.guanglin.pptGen.pptUtility.PPTFactory;
import lombok.NonNull;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.OptionBuilder;
import org.apache.commons.cli.Options;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;

import static com.guanglin.pptGen.Constants.APPCONFIG_PRO_NAME;
import static com.guanglin.pptGen.Constants.DATASOURCE_TYPE;
import static com.guanglin.pptGen.Constants.PROS;
import static com.guanglin.pptGen.Constants.PRO_CONFIG_PATH;
import static com.guanglin.pptGen.Constants.PRO_DATA_PATH;
import static com.guanglin.pptGen.Constants.PRO_IMG_OUTDORR_PATH;
import static com.guanglin.pptGen.Constants.PRO_IMG_PATH;
import static com.guanglin.pptGen.Constants.PRO_OUTPUT_PATH;
import static com.guanglin.pptGen.Constants.PRO_TEMPLATE_PATH;

public class Main {

    private static Logger LOGGER = LogManager.getLogger("Main");


    private final static String ARG_NAME_HELP = "help";
    private final static String PROJECTPATH = "projectPath";

    private static Options argOpts = new Options();

    {
        argOpts.addOption(ARG_NAME_HELP, "打印帮助信息。");
        argOpts.addOption(PROJECTPATH, "设置启动配置路径（必须）");
        argOpts.addOption(OptionBuilder.withArgName("projectPath=")
                .hasArgs(2)
                .withValueSeparator()
                .withDescription("use value for given property")
                .create());
    }

    public static void main(String[] args) {

        // String appConfigPath = "/Users/pengyao/Workspaces/personal/guanglin-new/configuration/app.config";

        try {

            LOGGER.debug("开始解析参数。");
            // parse the command line arguments
            CommandLineParser cmdParse = new DefaultParser();
            CommandLine cmdLine = cmdParse.parse(argOpts, args);

            // print command help
            if (cmdLine.hasOption(ARG_NAME_HELP)) {
                (new HelpFormatter()).printHelp("guanglin-new", argOpts);
                // return 0;
            }
            LOGGER.debug("解析参数完成。");
            String projectPath = args[0];

            // load app config into PROS
            if (!appInit(projectPath + PRO_CONFIG_PATH)) {
                LOGGER.fatal("读取项目配置失败，请检查项目配置文件。");
                // return -1;
            }

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
            PPTFactory.genPPT(project);
        } catch (Exception ex) {
            LOGGER.error("遇到错误：" + ex);
            ex.printStackTrace();
        }

    }

    private static boolean appInit(String appConfigPath) {

        try {
            // read config files
            PROS.load(new InputStreamReader(new FileInputStream(appConfigPath), "UTF-8"));
            return true;
        } catch (FileNotFoundException ex) {
            return false;

        } catch (IOException ex) {
            return false;
        }

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
