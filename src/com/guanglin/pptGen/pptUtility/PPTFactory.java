package com.guanglin.pptGen.pptUtility;


import com.google.common.base.Strings;
import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.exception.PPTInvildTemplateException;
import com.guanglin.pptGen.model.Project;

import java.io.FileInputStream;
import java.io.FileNotFoundException;


/**
 * Created by pengyao on 30/05/2017.
 */
public class PPTFactory {
    public static void genPPT(String pptTemplatePath, Project project) throws
            PPTException {

        if (Strings.isNullOrEmpty(pptTemplatePath)) {
            throw new PPTInvildTemplateException("模版地址的输入值为空值。");
        }


        try {
            FileInputStream templateStream = new FileInputStream(pptTemplatePath);

            HanlderBase handler;

            if (pptTemplatePath.contains(".ppt")) {
                // deal with ppt template
                handler = new PPTHandler(templateStream, project);

            } else if (pptTemplatePath.contains(".pptx")) {
                // deal with pptx template
                handler = new PPTXHandler(templateStream, project);
            } else {
                throw new PPTInvildTemplateException("模版的文件格式不正确，请使用.ppt 或者 .pptx 文件模版");
            }


        } catch (FileNotFoundException ex) {
            throw new PPTInvildTemplateException("找不到相应的模版文件。");
        }

    }
}
