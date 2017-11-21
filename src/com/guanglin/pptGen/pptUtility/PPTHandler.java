package com.guanglin.pptGen.pptUtility;

import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.model.PPTTemplate;
import com.guanglin.pptGen.model.Project;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * Created by pengyao on 30/05/2017.
 */
public class PPTHandler extends HanlderBase {

    public PPTHandler(FileInputStream templateStream, Project project) {
        super(templateStream, project);
        template = new PPTTemplate<String>();
    }

    @Override
    public void extractTemplate() throws PPTException {

    }

    @Override
    protected void validateTemplate() throws PPTException {

    }

    @Override
    protected void generateOutputSlide() throws PPTException {

    }

    @Override
    public void writePPTFileStream(FileOutputStream fileOutputStream) throws PPTException {

    }
}