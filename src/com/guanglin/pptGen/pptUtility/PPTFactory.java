package com.guanglin.pptGen.pptUtility;

import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.exception.PPTInvildTemplateException;
import com.guanglin.pptGen.model.Project;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;


/**
 * Created by pengyao on 30/05/2017.
 */
public class PPTFactory {

    private final static Logger LOGGER = LogManager.getLogger("PPTFactory");

    public static void genPPT(Project project) throws
            PPTException {

        try {
            FileInputStream templateStream = new FileInputStream(project.getprojectPPTTemplatePath());

            HanlderBase handler;

            if (project.getprojectPPTTemplatePath().contains(".ppt")) {
                // deal with ppt template
                handler = new PPTHandler(templateStream, project);

            } else if (project.getprojectPPTTemplatePath().contains(".pptx")) {
                // deal with pptx template
                handler = new PPTXHandler(templateStream, project);
            } else {
                throw new PPTInvildTemplateException("模版的文件格式不正确，请使用.ppt 或者 .pptx 文件模版");
            }

            final String outputFilePath = project.getProjectOutputPath() + "/" + project.getName() + ".ppt";
            File outputFile = new File(outputFilePath);
            if (outputFile.exists()) {
                outputFile.delete();
            }

            outputFile.createNewFile();
            FileOutputStream outputStream = new FileOutputStream(outputFile, false);
            handler.writePPTFileStream(outputStream);

            LOGGER.info("新生成的PPT文件：" + outputFilePath);

        } catch (Exception ex) {
            throw new PPTInvildTemplateException("找不到相应的模版文件。", ex);
        }

    }
}
