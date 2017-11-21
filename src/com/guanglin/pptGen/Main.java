package com.guanglin.pptGen;

import com.google.common.base.Strings;
import com.guanglin.pptGen.datasource.DataSourceFactory;
import com.guanglin.pptGen.exception.DataSourceException;
import com.guanglin.pptGen.model.Project;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.Options;
import org.apache.commons.cli.ParseException;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;

import static com.guanglin.pptGen.Constants.APPCONFIG_PRO_NAME;
import static com.guanglin.pptGen.Constants.ARG_NAME_APPCONFIG;
import static com.guanglin.pptGen.Constants.ARG_NAME_PROJECT;
import static com.guanglin.pptGen.Constants.ARG_NAME_TEMPLATE;
import static com.guanglin.pptGen.Constants.DATASOURCE_TYPE;
import static com.guanglin.pptGen.Constants.PROS;

public class Main {

    private static Logger LOGGER = LogManager.getLogger("Main");


    public final static String ARG_NAME_HELP = "help";
    public final static String ARG_NAME_APPCONFIG = "appConfig";
    public final static String ARG_NAME_TEMPLATE = "template";
    public final static String ARG_NAME_PROCONFIG = "proConfig";

    private static Options argOpts = new Options();

    {
        argOpts.addOption(ARG_NAME_HELP, "打印帮助信息。");
        argOpts.addOption(ARG_NAME_APPCONFIG, "设置启动配置路径（必须）");
        argOpts.addOption(ARG_NAME_TEMPLATE, "设置PPT模版文件路径（必须）");
        argOpts.addOption(ARG_NAME_PROCONFIG, "项目文件路径");
    }

    public static void main(String[] args) {

        String appConfigPath;
        try {
            LOGGER.info("Start to parse the arguments");
            // parse the command line arguments
            CommandLineParser cmdParse = new DefaultParser();
            CommandLine cmdLine = cmdParse.parse(argOpts, args);

            // print commnd help
            if(cmdLine.hasOption(ARG_NAME_HELP)) {
                (new HelpFormatter()).printHelp("guanglin-new", argOpts);
                return;
            }

            // check the required config;
            if(!cmdLine.hasOption(ARG_NAME_APPCONFIG) ) {
                throw new ParseException("缺少启动配置或项目配置，请检查");
            }
            appConfigPath = cmdLine.getOptionValue(ARG_NAME_APPCONFIG);

            LOGGER.info("End to parse the arguments");

        } catch (ParseException e) {
            LOGGER.fatal("Parse the error failed", e);
            return;
        }

        if (!appInit(appConfigPath)) {
            LOGGER.fatal("Config the app failed.");
        }


        try {
            Project project = new Project();
            project.setOwner(PROS.getProperty(APPCONFIG_PRO_NAME));

            String datasourceType = (String) PROS.get(DATASOURCE_TYPE);
            project = DataSourceFactory.loadProjectData(datasourceType, project);

        } catch (InvalidFormatException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (DataSourceException e) {
            e.printStackTrace();
        }

        // read excel


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


}
