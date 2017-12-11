package com.guanglin.pptGen.pptUtility;


import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.exception.PPTInvildTemplateException;
import com.guanglin.pptGen.model.Project;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;


/**
 * Created by pengyao on 30/05/2017.
 */
public class PPTFactory {
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

            File outputFile = new File(project.getProjectOutputPath() + project.getName() + ".ppt");
            if (!outputFile.exists()) {
                outputFile.createNewFile();
            }

            FileOutputStream outputStream = new FileOutputStream(outputFile, false);
            handler.writePPTFileStream(outputStream);


        } catch (FileNotFoundException ex) {
            throw new PPTInvildTemplateException("找不到相应的模版文件。");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
