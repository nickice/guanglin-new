package com.guanglin.pptGen;

import com.guanglin.pptGen.exception.ProjectValidationException;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.Project;
import com.guanglin.pptGen.pptUtility.PPTFactory;
import com.guanglin.pptGen.project.ProjectBuilder;
import org.apache.commons.cli.CommandLine;
import org.apache.commons.cli.CommandLineParser;
import org.apache.commons.cli.DefaultParser;
import org.apache.commons.cli.HelpFormatter;
import org.apache.commons.cli.OptionBuilder;
import org.apache.commons.cli.Options;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;

import static com.guanglin.pptGen.Constants.PROS;
import static com.guanglin.pptGen.Constants.PRO_CONFIG_PATH;

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

            Project project = ProjectBuilder.build(projectPath);
            PPTFactory.genPPT(project);
        } catch (ProjectValidationException pvEx) {
            if (pvEx.getInvalidItems() != null && pvEx.getInvalidItems().size() > 0) {
                for (Item item : pvEx.getInvalidItems().keySet()) {
                    LOGGER.error(pvEx.getInvalidItems().get(item) + ": " + item.getDescription());
                }
            }
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

}
