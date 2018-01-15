package com.guanglin.pptGen.pptUtility;

import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.PPTTemplate;
import com.guanglin.pptGen.model.Project;
import com.guanglin.pptGen.utils.ObjectUtils;
import org.apache.log4j.LogManager;
import org.apache.log4j.Logger;
import org.apache.poi.hslf.usermodel.HSLFAutoShape;
import org.apache.poi.hslf.usermodel.HSLFFill;
import org.apache.poi.hslf.usermodel.HSLFPictureData;
import org.apache.poi.hslf.usermodel.HSLFPictureShape;
import org.apache.poi.hslf.usermodel.HSLFShape;
import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;
import org.apache.poi.hslf.usermodel.HSLFSlideShowImpl;
import org.apache.poi.hslf.usermodel.HSLFTable;
import org.apache.poi.hslf.usermodel.HSLFTableCell;
import org.apache.poi.hslf.usermodel.HSLFTextBox;
import org.apache.poi.hslf.usermodel.HSLFTextParagraph;
import org.apache.poi.hslf.usermodel.HSLFTextRun;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TableCell;

import java.awt.*;
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
    private String fontFamily = null;

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

        LOGGER.debug("start to validate PPT template");
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

        /*
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
        */
        LOGGER.debug("end to validate PPT template");
    }

    @Override
    protected void generateOutputSlide() throws PPTException {

        LOGGER.info("****  开始生成项目ppt文件  ****");
        LOGGER.debug("start to generate output slide");
        generatedShow = new HSLFSlideShow();
        generatedShow.setPageSize(templateShow.getPageSize());


        try {

            // foreach the items
            for (Item item : this.project.getItems()) {

                int slideCount = 0;
                int insideCaptureCount = item.getInsideCaptures().size();

                // if the image less than 3
                if (item.getInsideCaptures().size() % 3 == 0) {
                    //  template 1
                    LOGGER.debug("use 2 images template to deal with the outdoor image.");
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(1), item, true);
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(1), item, false);
                    slideCount = slideCount + 2;
                } else if (item.getInsideCaptures().size() % 3 == 1) {
                    // template 1
                    LOGGER.debug("use 2 images template to deal with the outdoor image.");
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(1), item, true);
                    slideCount++;
                } else if (item.getInsideCaptures().size() % 3 == 2) {
                    // template 2
                    LOGGER.debug("use 2 image template to deal with the outdoor image.");
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(2), item, true);
                    slideCount++;
                }

                LOGGER.debug("start to use 3 images template to deal with item indoor captures, size is " + item.getInsideCaptures().size());
                while (!item.getInsideCaptures().empty()) {
                    // template 3
                    buildSlide((HSLFSlide) this.template.getTemplateSlideMap().get(2), item, false);
                    slideCount++;
                }

                LOGGER.info("社区: " + item.getDescription() + ";");
                LOGGER.info("--生成幻灯片：" + slideCount + "张;");
                LOGGER.info("--内景图片数量: " + insideCaptureCount);
            }

        } catch (Exception ex) {
            LOGGER.error(ex);
            throw new PPTException("致命错误：生成PPT时失败了。", ex);
        }

    }


    private void buildSlide(HSLFSlide templateSlide, Item item, boolean setOutDoor) throws Exception {

        // create a slide
        HSLFSlide slide = generatedShow.createSlide();


        //region set background
        if (templateSlide.getBackground().getShapeType() == ShapeType.RECT) {

            // TODO: investigate how to set background
            LOGGER.debug("start to set background.");
            slide.setFollowMasterBackground(false);
            if (templateSlide.getBackground() != null &&
                    templateSlide.getBackground().getFill() != null &&
                    templateSlide.getBackground().getFill().getPictureData() != null) {
                byte[] templateBackgroundData = templateSlide.getBackground().getFill().getPictureData().getRawData();
                PictureData.PictureType pictureType = templateSlide.getBackground().getFill().getPictureData().getType();
                HSLFPictureData pictureData = generatedShow.addPicture(templateBackgroundData, pictureType);
                slide.getBackground().getFill().setFillType(HSLFFill.FILL_PICTURE);
                slide.getBackground().getFill().setPictureData(pictureData);
            }
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
                HSLFTextBox templateTextBox = (HSLFTextBox) shape;
                String originalText = templateTextBox.getText();
                HSLFTextBox newTextBox = new HSLFTextBox();
                newTextBox.setText(replacedItemText(item, originalText));
                newTextBox.setAnchor(templateTextBox.getAnchor());

                for (int i = 0; i < newTextBox.getTextParagraphs().size(); i++) {
                    HSLFTextParagraph textParagraph = newTextBox.getTextParagraphs().get(i);
                    HSLFTextParagraph templateParagragh = templateTextBox.getTextParagraphs().get(i);

                    textParagraph.setParagraphStyle(templateParagragh.getParagraphStyle());

                    // indent setting
                    textParagraph.setIndent(templateParagragh.getIndent());
                    textParagraph.setIndentLevel(templateParagragh.getIndentLevel());

                    // bullet setting
                    ObjectUtils.copyPropertyByName("BulletStyle", templateParagragh, textParagraph);
                    ObjectUtils.copyPropertyByName("BulletColor", templateParagragh, textParagraph);
                    ObjectUtils.copyPropertyByName("BulletChar", templateParagragh, textParagraph);
                    ObjectUtils.copyPropertyByName("BulletSize", templateParagragh, textParagraph);

                    for (int n = 0; n < textParagraph.getTextRuns().size(); n++) {
                        textParagraph.getTextRuns().get(n).setFontFamily(templateParagragh.getTextRuns().get(n).getFontFamily());
                        textParagraph.getTextRuns().get(n).setFontSize(templateParagragh.getTextRuns().get(n).getFontSize());
                        textParagraph.getTextRuns().get(n).setBold(templateParagragh.getTextRuns().get(n).isBold());
                        textParagraph.getTextRuns().get(n).setFontColor(templateParagragh.getTextRuns().get(n).getFontColor());
                        textParagraph.getTextRuns().get(n).setCharacterStyle(templateParagragh.getTextRuns().get(n).getCharacterStyle());
                    }

                    // print debug information
/*
                    for (HSLFTextRun textRun : textParagraph.getTextRuns()) {
                        if (textRun != null) {
                            LOGGER.debug("font size: " + textRun.getFontSize());
                            LOGGER.debug("font family: " + textRun.getFontFamily());
                            LOGGER.debug("raw text： " + textRun.getRawText());
                            LOGGER.debug("text cap" + textRun.getTextCap());
                        }
                    }
                    */
                }

                slide.addShape(newTextBox);
            } else if (shape instanceof HSLFTable) {
                HSLFTable templateTable = (HSLFTable) shape;
                HSLFTable newTable = slide.createTable(templateTable.getNumberOfRows(), templateTable.getNumberOfColumns());
                newTable.setAnchor(templateTable.getAnchor());
                newTable.setShapeId(templateTable.getShapeId());

                // loop the row and column
                if (templateTable.getNumberOfRows() > 0 && templateTable.getNumberOfColumns() > 0) {
                    for (int r = 0; r < templateTable.getNumberOfRows(); r++) {
                        newTable.setRowHeight(r, templateTable.getRowHeight(r));
                        for (int c = 0; c < templateTable.getNumberOfColumns(); c++) {
                            // set the column style if the row index is 0
                            if (r == 0) {
                                newTable.setColumnWidth(c, templateTable.getColumnWidth(c));
                            }

                            if (templateTable.getCell(r, c) != null) {

                                HSLFTableCell templateCell = templateTable.getCell(r, c);

                                // replace content
                                ObjectUtils.copyPropertyByName("FillColor", templateCell, newTable.getCell(r, c));
                                newTable.getCell(r, c).setText(replacedItemText(item, templateCell.getText()));

                                // set content style
                                for (int n = 0; n < newTable.getCell(r, c).getTextParagraphs().size(); n++) {
                                    if (templateCell.getTextParagraphs().get(n) != null && newTable.getCell(r, c).getTextParagraphs().get(n) != null) {
                                        for (int m = 0; m < newTable.getCell(r, c).getTextParagraphs().get(n).getTextRuns().size(); m++) {
                                            HSLFTextRun destTextRun = newTable.getCell(r, c).getTextParagraphs().get(n).getTextRuns().get(m);
                                            HSLFTextRun sourceTextRun = templateCell.getTextParagraphs().get(n).getTextRuns().get(m);

                                            destTextRun.setCharacterStyle(sourceTextRun.getCharacterStyle());
                                            ObjectUtils.copyPropertyByName("FontFamily", sourceTextRun, destTextRun);
                                            ObjectUtils.copyPropertyByName("FontSize", sourceTextRun, destTextRun);
                                            destTextRun.setFontColor(sourceTextRun.getFontColor());
                                            destTextRun.setBold(sourceTextRun.isBold());

                                        }
                                    }
                                }

                                newTable.getCell(r, c).setLineColor(templateCell.getLineColor());
                                newTable.getCell(r, c).setLineCompound(templateCell.getLineCompound());
                                newTable.getCell(r, c).setLineDash(templateCell.getLineDash());
                                newTable.getCell(r, c).setLineCap(templateCell.getLineCap());
                                newTable.getCell(r, c).setLineBackgroundColor(templateCell.getLineBackgroundColor());
                                newTable.getCell(r, c).setVerticalAlignment(templateCell.getVerticalAlignment());
                                newTable.getCell(r, c).setHorizontalCentered(true);

                                // set the cell border style
                                for (TableCell.BorderEdge borderEdge : TableCell.BorderEdge.values()) {
                                    if (templateCell.getBorderColor(borderEdge) != null) {
                                        newTable.getCell(r, c).setBorderColor(borderEdge, templateCell.getBorderColor(borderEdge));
                                    } else {
                                        newTable.getCell(r, c).setBorderColor(borderEdge, Color.black);
                                    }

                                    if (templateCell.getBorderCompound(borderEdge) != null) {
                                        newTable.getCell(r, c).setBorderCompound(borderEdge, templateCell.getBorderCompound(borderEdge));
                                    }

                                    if (templateCell.getBorderStyle(borderEdge) != null) {
                                        newTable.getCell(r, c).setBorderStyle(borderEdge, templateCell.getBorderStyle(borderEdge));
                                    }

                                    if (templateCell.getBorderDash(borderEdge) != null) {
                                        newTable.getCell(r, c).setBorderDash(borderEdge, templateCell.getBorderDash(borderEdge));
                                    }

                                    if (templateCell.getBorderWidth(borderEdge) != null) {
                                        newTable.getCell(r, c).setBorderWidth(borderEdge, templateCell.getBorderWidth(borderEdge));
                                    }
                                }

                                // set anchor
                                newTable.getCell(r, c).setAnchor(templateCell.getAnchor());
                            }
                        }
                    }

                }
                slide.addShape(newTable);


            } else if (shape instanceof HSLFAutoShape) {
                slide.addShape((HSLFAutoShape) shape);

            } else {
                LOGGER.error("致命错误：第一幻灯片中存在一个未知的图形: " + shape.getClass());
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