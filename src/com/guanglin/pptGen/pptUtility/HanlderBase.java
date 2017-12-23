package com.guanglin.pptGen.pptUtility;

import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.PPTTemplate;
import com.guanglin.pptGen.model.Project;
import lombok.NonNull;

import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * This is base class of pptHandler
 * There are two inherit classes: pptHandler, pptxHandler
 * <p>
 * Created by pengyao on 30/05/2017.
 */
public abstract class HanlderBase<S, P> {

    protected FileInputStream templateFileStream;
    protected Project project;
    protected PPTTemplate template;

    public HanlderBase(FileInputStream templateFileStream, Project project) {
        this.templateFileStream = templateFileStream;
        this.project = project;
    }

    protected abstract void extractTemplate() throws PPTException;

    protected abstract void validateTemplate() throws PPTException;

    protected abstract void generateOutputSlide() throws PPTException;

    protected String replacedItemText(final Item item, @NonNull String textTemplate) {

        String text = textTemplate;
        for (String fieldName : item.getFields().keySet()) {
            if (textTemplate.contains("[" + fieldName + "]")) {
                text = text.replace("[" + fieldName + "]", item.getFields().get(fieldName));
            }
        }

        return text;
    }

    public abstract void writePPTFileStream(FileOutputStream fileOutputStream) throws PPTException;

}
