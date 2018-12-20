package com.guanglin.pptGen.pptUtility;

import com.guanglin.pptGen.exception.PPTException;
import com.guanglin.pptGen.model.Item;
import com.guanglin.pptGen.model.PPTTemplate;
import com.guanglin.pptGen.model.Project;
import com.guanglin.pptGen.utils.ImageUtils;
import com.guanglin.pptGen.utils.ObjectUtils;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TableCell;
import org.apache.poi.xslf.usermodel.SlideLayout;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTableCell;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;

import java.awt.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by pengyao on 30/05/2017.
 */
public class PPTXHandler<S, P> extends HanlderBase {

    private static Logger LOGGER = LogManager.getLogger("PPTXHandler");

    private XMLSlideShow templateShow;
    private XMLSlideShow generatedShow;

    public PPTXHandler(FileInputStream stream, Project project) throws PPTException {
        super(stream, project);
        template = new PPTTemplate<XSLFSlide>();

        try {
            templateShow = new XMLSlideShow(stream);
        } catch (IOException ex) {
            LOGGER.fatal(PPTLogMessage.FATAL_TEMPLATE_MISS);
            throw new PPTException(PPTLogMessage.FATAL_TEMPLATE_MISS);
        }

        this.extractTemplate();
    }

    @Override
    protected void extractTemplate() throws PPTException {

        this.validateTemplate();

        // init template instance
        this.template = new PPTTemplate();

        Map<Integer, XSLFSlide> slideMap = new HashMap<>();
        for (XSLFSlide slide : this.templateShow.getSlides()) {
            slideMap.put(slide.getSlideNumber(), slide);
        }

        this.template.setTemplateSlideMap(slideMap);

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

        /*
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
*/
    }

    @Override
    protected void generateOutputSlide() throws PPTException {

        generatedShow = new XMLSlideShow();
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
                    buildSlide((XSLFSlide) this.template.getTemplateSlideMap().get(1), item, true);
                    buildSlide((XSLFSlide) this.template.getTemplateSlideMap().get(1), item, false);
                    slideCount = slideCount + 2;
                } else if (item.getInsideCaptures().size() % 3 == 1) {
                    // template 1
                    LOGGER.debug("use 2 images template to deal with the outdoor image.");
                    buildSlide((XSLFSlide) this.template.getTemplateSlideMap().get(1), item, true);
                    slideCount++;
                } else if (item.getInsideCaptures().size() % 3 == 2) {
                    // template 2
                    LOGGER.debug("use 2 image template to deal with the outdoor image.");
                    buildSlide((XSLFSlide) this.template.getTemplateSlideMap().get(2), item, true);
                    slideCount++;
                }

                LOGGER.debug("start to use 3 images template to deal with item indoor captures, size is " + item.getInsideCaptures().size());
                while (!item.getInsideCaptures().empty()) {
                    // template 3
                    buildSlide((XSLFSlide) this.template.getTemplateSlideMap().get(2), item, false);
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

    private void buildSlide(XSLFSlide templateSlide, Item item, boolean setOutDoor) throws Exception {

        // create a slide
        XSLFSlide slide = generatedShow.createSlide();


        //region set background
        if (templateSlide.getBackground().getShapeType() == ShapeType.RECT) {

            // TODO: investigate how to set background

            /*
            LOGGER.debug("start to set background.");
            slide.setFollowMasterBackground(false);
            if (templateSlide.getBackground() != null &&
                    templateSlide.getBackground() != null &&
                    templateSlide.getBackground().getFill().getPictureData() != null) {
                byte[] templateBackgroundData = templateSlide.getBackground().getFill().getPictureData().getRawData();
                PictureData.PictureType pictureType = templateSlide.getBackground().getFill().getPictureData().getType();
                XSLFPictureData pictureData = generatedShow.addPicture(templateBackgroundData, pictureType);
                slide.getBackground().getFill().setFillType(XSLFFill.FILL_PICTURE);
                slide.getBackground().getFill().setPictureData(pictureData);
            }
            */
        }
        //endregion

        boolean hasSetOutdoor = false;
        for (XSLFShape shape : templateSlide.getShapes()) {
            if (shape instanceof XSLFPictureShape) {
                XSLFPictureData pictureData;
                XSLFPictureShape pictureShape;

                if (setOutDoor && !hasSetOutdoor) {
                    // set the outdoor image
                    LOGGER.debug("start to build outdoor image shape ");
                    pictureData = generatedShow.addPicture(ImageUtils.getImageBytes(item.getOutsideCapture().getCapturePath()), PictureData.PictureType.JPEG);
                    hasSetOutdoor = true;
                } else {
                    LOGGER.debug("start to build indoor image shape ");
                    pictureData = generatedShow.addPicture(ImageUtils.getImageBytes(item.getInsideCaptures().pop().getCapturePath()), PictureData.PictureType.JPEG);
                }
                //pictureShape = new XSLFPictureShape(pictureData);
                //pictureShape.setAnchor(shape.getAnchor());
                // slide.addShape(pictureShape);

            } else if (shape instanceof XSLFTextBox) {
                // set the description
                LOGGER.debug("start to build text shape ");
                XSLFTextBox templateTextBox = (XSLFTextBox) shape;
                String originalText = templateTextBox.getText();
                XSLFTextBox newTextBox = null;
                newTextBox.setText(replacedItemText(item, originalText));
                newTextBox.setAnchor(templateTextBox.getAnchor());

                for (int i = 0; i < newTextBox.getTextParagraphs().size(); i++) {
                    XSLFTextParagraph textParagraph = newTextBox.getTextParagraphs().get(i);
                    XSLFTextParagraph templateParagragh = templateTextBox.getTextParagraphs().get(i);

                  //  textParagraph.setParagraphStyle(templateParagragh.getParagraphStyle());

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
                  //      textParagraph.getTextRuns().get(n).setCharacterStyle(templateParagragh.getTextRuns().get(n).getCharacterStyle());
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
            } else if (shape instanceof XSLFTable) {
                XSLFTable templateTable = (XSLFTable) shape;
                XSLFTable newTable = slide.createTable(templateTable.getNumberOfRows(), templateTable.getNumberOfColumns());
                newTable.setAnchor(templateTable.getAnchor());
              //  newTable.setShapeId(templateTable.getShapeId());

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


                                XSLFTableCell templateCell = templateTable.getCell(r, c);

                                // replace content
                                ObjectUtils.copyPropertyByName("FillColor", templateCell, newTable.getCell(r, c));
                                newTable.getCell(r, c).setText(replacedItemText(item, templateCell.getText()));

                                // set content style
                                for (int n = 0; n < newTable.getCell(r, c).getTextParagraphs().size(); n++) {
                                    if (templateCell.getTextParagraphs().get(n) != null && newTable.getCell(r, c).getTextParagraphs().get(n) != null) {
                                        for (int m = 0; m < newTable.getCell(r, c).getTextParagraphs().get(n).getTextRuns().size(); m++) {
                                            XSLFTextRun destTextRun = newTable.getCell(r, c).getTextParagraphs().get(n).getTextRuns().get(m);
                                            XSLFTextRun sourceTextRun = templateCell.getTextParagraphs().get(n).getTextRuns().get(m);

                                       //     destTextRun.setCharacterStyle(sourceTextRun.getCharacterStyle());
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
                             //   newTable.getCell(r, c).setLineBackgroundColor(templateCell.getLineBackgroundColor());
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


            } else if (shape instanceof XSLFAutoShape) {
                slide.addShape((XSLFAutoShape) shape);

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
