package com.guanglin.pptGen.pptUtility;

import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.PPTTemplate;
import com.guanglin.pptGen.model.Project;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayDeque;
import java.util.Queue;

/**
 * Created by pengyao on 30/05/2017.
 */
public class PPTXHandler extends HanlderBase {

    private static Logger LOGGER = LogManager.getLogger("PPTXHandler");

    private XMLSlideShow templateShow;
    private XMLSlideShow generatedShow;

    public PPTXHandler(FileInputStream stream, Project project) throws PPTException {
        super(stream, project);
        template = new PPTTemplate<XSLFSlide>();

        try {
            templateShow = new XMLSlideShow(stream);
        } catch (IOException ex) {
            LOGGER.fatal("致命错误：找不到PPT模版文件。");
            throw new PPTException("致命错误：找不到PPT模版文件。");
        }

        this.extractTemplate();
    }

    @Override
    protected void extractTemplate() throws PPTException {

        this.validateTemplate();

        // init template instance
        this.template = new PPTTemplate();

        Queue<XSLFSlide> slideQueue = new ArrayDeque<XSLFSlide>();
        for (XSLFSlide slide : this.templateShow.getSlides()) {
            slideQueue.add(slide);
        }

        this.template.setSlidesQueue(slideQueue);

    }

    @Override
    protected void validateTemplate() throws PPTException {
        // 1. the slide count should be greater than 1
        if (this.templateShow.getSlides() == null ||
                this.templateShow.getSlides().size() < 2) {
            throw new PPTException("致命错误：这个模版没有任何幻灯片，或者只有一个幻灯片。");
        }

        // 2. the first slide must have two image sharps.
        XSLFSlide firstSlide = this.templateShow.getSlides().get(0);
        if (firstSlide == null
                || firstSlide.getShapes() == null
                || firstSlide.getShapes().size() < 2) {
            throw new PPTException("致命错误：这个模版第一个幻灯片没有要求的图片样式。");
        }

        // 3. each slide must have a text input sharp
        for (int i = 1; i < this.templateShow.getSlides().size(); i++) {

            int slideIndex = i + 1;
            XSLFSlide slide = this.templateShow.getSlides().get(i);
            if (slide == null
                    || slide.getShapes() == null
                    || slide.getShapes().size() < 2) {
                throw new PPTException(String.format("致命错误：这个模版第%i模版图片样式不符合要求。", slideIndex));
            }

            boolean hasTextShape = false;
            for (XSLFShape shape : slide.getShapes()) {
                // indicate the shape is textbox or not
                if (shape instanceof XSLFTextBox) {
                    hasTextShape = true;
                    break;
                }
            }

            // throw exception to tell there is a slick doesn't have a textbox
            if (!hasTextShape) {
                throw new PPTException(String.format("致命错误：这个模版第%i幻灯片不包含文本输入", slideIndex));
            }
        }

        // TODO: Check the each slide has image shapes.

    }

    @Override
    protected void generateOutputSlide() throws PPTException {

        generatedShow = new XMLSlideShow();

        try {
            XSLFSlide firstSlideTemplate = (XSLFSlide) this.template.getSlidesQueue().poll();

            // foreach the items
            for (Item item : this.project.getItems()) {
                XSLFSlide firstSlide = generatedShow.createSlide();

                // 1. the first slide
                for (XSLFShape shape : firstSlideTemplate.getShapes()) {
                    if (shape instanceof XSLFPictureShape) {

                        if (shape.getShapeName().trim() == "外景") {
                            // 1.1 set the outdoor image
                            ((XSLFPictureShape) shape).getPictureData().setData(item.getOutsideCapture().getBytes());
                        } else {
                            // 1.2 set the indoor image
                            ((XSLFPictureShape) shape).getPictureData().setData(item.getInsideCaptures().pop().getBytes());
                        }
                    } else if (shape instanceof XSLFTextBox) {
                        // 1.3 set the description
                        ((XSLFTextBox) shape).setText(item.getDescription());

                    } else {
                        throw new PPTException("致命错误：第一幻灯片中存在一个未知的图形");
                    }

                    firstSlide.addShape(shape);
                }


                // 2. calculate the other slides
                for (Object slideTemplateObj : this.template.getSlidesQueue()) {

                    XSLFSlide slide = generatedShow.createSlide();
                    XSLFSlide slideTemplate = (XSLFSlide) slideTemplateObj;

                    // foreach shape in the slide
                    for (XSLFShape shape : slideTemplate) {
                        if (shape instanceof XSLFPictureShape) {

                            // 2.2 set the indoor image if the sharp is image
                            ((XSLFPictureShape) shape).getPictureData().setData(item.getInsideCaptures().pop().getBytes());
                        } else if (shape instanceof XSLFTextBox) {
                            // 1.3 set the description
                            ((XSLFTextBox) shape).setText(item.getDescription());
                        } else {
                            throw new PPTException("致命错误：普通幻灯片中存在一个未知的图形");
                        }
                        slide.addShape(shape);
                    }

                }
            }
        } catch (IOException exception) {
            throw new PPTException("致命错误：构造PPTX时失败了。");
        }
    }

    @Override
    public void writePPTFileStream(FileOutputStream fileOutputStream) throws PPTException {
        try {
            generateOutputSlide();
            this.generatedShow.write(fileOutputStream);
        } catch (IOException ex) {
            throw new PPTException("致命错误");
        }

    }
}
