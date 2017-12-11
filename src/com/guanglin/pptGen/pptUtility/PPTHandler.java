package com.guanglin.pptGen.pptUtility;

import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.PPTTemplate;
import com.guanglin.pptGen.model.Project;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.hslf.usermodel.HSLFFill;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by pengyao on 30/05/2017.
 */
public class PPTHandler extends HanlderBase {

    private static Logger LOGGER = LogManager.getLogger(PPTHandler.class);

    private HSLFSlideShow templateShow = null;
    private HSLFSlideShow generatedShow = null;

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

        // init template instance
        this.template = new PPTTemplate();

        // foreach template slides
        Map<Integer, HSLFSlide> slideMap = new HashMap<>();
        for (HSLFSlide slide : this.templateShow.getSlides()) {
            slideMap.put(slide.getSlideNumber(), slide);
        }

        this.template.setTemplateSlideMap(slideMap);

    }

    @Override
    protected void validateTemplate() throws PPTException {

        LOGGER.info("start to validate PPT template");
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
        LOGGER.info("end to validate PPT template");
    }

    @Override
    protected void generateOutputSlide() throws PPTException {

        LOGGER.info("start to generate output slide");
        generatedShow = new HSLFSlideShow();

        try {

            // foreach the items
            for (Item item : this.project.getItems()) {

                LOGGER.info("start to deal with item: " + item.getDescription());
                LOGGER.debug("item name: " + item.getDescription());
                LOGGER.debug("item capture size: " + item.getInsideCaptures().size());

                // if the image less than 3
                if (item.getInsideCaptures().size() % 3 == 0) {
                    //  template 1
                    LOGGER.debug("use 2 images template to deal with the outdoor image.");
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(2), item, true);
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(2), item, false);
                } else if (item.getInsideCaptures().size() % 3 == 1) {
                    // template 2
                    LOGGER.debug("use 2 images template to deal with the outdoor image.");
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(2), item, true);
                } else if (item.getInsideCaptures().size() % 3 == 2) {
                    // template 1,2
                    LOGGER.debug("use 1 image template to deal with the outdoor image.");
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(1), item, true);
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(2), item, false);
                }

                LOGGER.debug("start to use 3 images template to deal with item indoor captures, size is " + item.getInsideCaptures().size());
                while (!item.getInsideCaptures().empty()) {
                    // template 3
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(3), item, false);
                }
            }

        } catch (Exception ex) {
            LOGGER.error(ex);
            throw new PPTException("致命错误：构造PPTX时失败了。");
        }

    }


    private void buildSlide(HSLFSlide templateSlide, Item item, boolean setOutDoor) throws PPTException, IOException {

        // create a slide
        HSLFSlide slide = generatedShow.createSlide();

        //region set background
        if (templateSlide.getBackground().getShapeType() == ShapeType.RECT) {

            LOGGER.debug("start to set background.");
            slide.setFollowMasterBackground(false);
            byte[] templateBackgroundData = templateSlide.getBackground().getFill().getPictureData().getRawData();
            PictureData.PictureType pictureType = templateSlide.getBackground().getFill().getPictureData().getType();
            HSLFPictureData pictureData = generatedShow.addPicture(templateBackgroundData, pictureType);
            slide.getBackground().getFill().setFillType(HSLFFill.FILL_PATTERN);
            slide.getBackground().getFill().setPictureData(pictureData);

        }
        //endregion

        boolean hasSetOutdoor = false;
        for (HSLFShape shape : templateSlide.getShapes()) {
            if (shape instanceof HSLFPictureShape) {
                HSLFPictureData pictureData;
                HSLFPictureShape pictureShape;

                if (setOutDoor && !hasSetOutdoor) {
                    // set the outdoor image
                    LOGGER.debug("start to build outdoor image shape ");
                    pictureData = generatedShow.addPicture(item.getOutsideCapture().getBytes(), PictureData.PictureType.JPEG);
                    hasSetOutdoor = true;
                } else {
                    LOGGER.debug("start to build indoor image shape ");
                    pictureData = generatedShow.addPicture(item.getInsideCaptures().pop().getBytes(), PictureData.PictureType.JPEG);
                }
                pictureShape = new HSLFPictureShape(pictureData);
                pictureShape.setAnchor(shape.getAnchor());
                slide.addShape(pictureShape);

            } else if (shape instanceof HSLFTextBox) {
                // set the description
                LOGGER.debug("start to build text shape ");
                HSLFTextBox templateTextBox = ((HSLFTextBox) shape);
                String originalText = templateTextBox.getText();
                //TODO: need to test on chinese ppt
                templateTextBox.setText(replacedItemText(item, originalText));
                slide.addShape(templateTextBox);

            } else {
                throw new PPTException("致命错误：第一幻灯片中存在一个未知的图形");
            }

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