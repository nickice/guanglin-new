package com.guanglin.pptGen.pptUtility;

import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.model.PPTTemplate;
import com.guanglin.pptGen.model.Project;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTextBox;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

/**
 * Created by pengyao on 30/05/2017.
 */
public class PPTHandler extends HanlderBase {

    private HSLFSlideShow templateShow = null;

    public PPTHandler(FileInputStream templateStream, Project project) {
        super(templateStream, project);
        template = new PPTTemplate<HSLFSlideShow>();
        try {
            templateShow = new HSLFSlideShow(new HSLFSlideShowImpl(templateStream));
            this.extractTemplate();
        } catch (IOException e) {
            e.printStackTrace();
        } catch (PPTException e) {
            e.printStackTrace();
        }

    }

    @Override
    public void extractTemplate() throws PPTException {
        this.validateTemplate();

    }

    @Override
    protected void validateTemplate() throws PPTException {
        // 1. the slide count should be greater than 1
        if (this.templateShow.getSlides() == null ||
                this.templateShow.getSlides().size() < 2) {
            throw new PPTException("致命错误：这个模版没有任何幻灯片，或者只有一个幻灯片。");
        }

        // 2. the first slide must have two image sharps.
        HSLFSlide firstSlide = this.templateShow.getSlides().get(0);
        if (firstSlide == null
                || firstSlide.getShapes() == null
                || firstSlide.getShapes().size() < 2) {
            throw new PPTException("致命错误：这个模版第一个幻灯片没有要求的图片样式。");
        }

        // 3. each slide must have a text input sharp
        for (int i = 1; i < this.templateShow.getSlides().size(); i++) {

            int slideIndex = i + 1;
            HSLFSlide slide = this.templateShow.getSlides().get(i);
            if (slide == null
                    || slide.getShapes() == null
                    || slide.getShapes().size() < 2) {
                throw new PPTException(String.format("致命错误：这个模版第%i模版图片样式不符合要求。", slideIndex));
            }

            boolean hasTextShape = false;
            for (HSLFShape shape : slide.getShapes()) {
                // indicate the shape is textbox or not
                if (shape instanceof HSLFTextBox) {
                    hasTextShape = true;
                    break;
                }
            }

            // throw exception to tell there is a slick doesn't have a textbox
            if (!hasTextShape) {
                throw new PPTException(String.format("致命错误：这个模版第%i幻灯片不包含文本输入", slideIndex));
            }
        }
    }

    @Override
    protected void generateOutputSlide() throws PPTException {

    }

    @Override
    public void writePPTFileStream(FileOutputStream fileOutputStream) throws PPTException {

    }
}