package com.j2ooxml.pptx.util;

import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.imaging.ImageInfo;
import org.apache.commons.imaging.ImageReadException;
import org.apache.commons.imaging.Imaging;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualPictureProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTApplicationNonVisualDrawingProps;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackground;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPictureNonVisual;
import org.w3c.dom.Node;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleRule;
import org.w3c.dom.css.CSSStyleSheet;

public class PptxUtil {

    public static void removeSlideBackground(XSLFSlide slide) {
        XSLFBackground bg = slide.getBackground();
        CTBackground xmlBg = (CTBackground) bg.getXmlObject();
        Node bgDomNode = xmlBg.getDomNode();
        bgDomNode.getParentNode().removeChild(bgDomNode);
    }

    public static Double getDefaultFontSize(XSLFTextShape textShape) {
        List<XSLFTextParagraph> textParagraphs = textShape.getTextParagraphs();
        Double defaultFontSize = null;
        if (CollectionUtils.isNotEmpty(textParagraphs)) {
            XSLFTextParagraph textParagraph = textParagraphs.get(0);
            defaultFontSize = textParagraph.getDefaultFontSize();
        }
        return defaultFontSize;
    }

    public static void embedPicture(XSLFPictureShape picture, Path picturePath) throws IOException, ImageReadException {
        byte[] pictureBytes = IOUtils.toByteArray(Files.newInputStream(picturePath));
        picture.getPictureData().setData(pictureBytes);
        CTPicture xmlPicture = (CTPicture) picture.getXmlObject();
        CTPictureNonVisual nvPicPr = xmlPicture.getNvPicPr();
        CTNonVisualPictureProperties cNvPicPr = nvPicPr.getCNvPicPr();
        boolean verticalCenter = !cNvPicPr.isSetPreferRelativeResize() || cNvPicPr.getPreferRelativeResize();
        boolean smartStretch = !cNvPicPr.getPicLocks().getNoChangeAspect();
        if (smartStretch || verticalCenter) {
            Rectangle2D anchor = picture.getAnchor();
            double x = anchor.getX();
            double y = anchor.getY();

            double wp = anchor.getWidth();
            double hp = anchor.getHeight();

            ImageInfo imageInfo = Imaging.getImageInfo(picturePath.toFile());
            int physicalHeightDpi = imageInfo.getPhysicalHeightDpi();
            if (physicalHeightDpi < 0) {
                physicalHeightDpi = 72;
            }
            int physicalWidthDpi = imageInfo.getPhysicalWidthDpi();
            if (physicalWidthDpi < 0) {
                physicalWidthDpi = 72;
            }
            double wi = Math.round(72. * imageInfo.getWidth() / physicalHeightDpi);
            double hi = Math.round(72. * imageInfo.getHeight() / physicalWidthDpi);

            double w = wp;
            double h = hp;
            double dx = 0;
            double dy = 0;
            if (verticalCenter) {
                w = wi;
                h = hi;
                dx = wp - wi;
                dy = (hp - hi) / 2;
            } else if (smartStretch) {
                if (wp / hp > wi / hi) {
                    w = wi * hp / hi;
                    dx = (wp - w) / 2;
                } else if (wp / hp < wi / hi) {
                    h = hi * wp / wi;
                    dy = (hp - h) / 2;
                }
            }
            anchor.setRect(x + dx, y + dy, w, h);
            picture.setAnchor(anchor);
        }
    }

    public static void applyPictureCss(XSLFPictureShape picture, CSSStyleSheet css, String name, Object value) {
        CSSRuleList rules = css.getCssRules();
        if (value instanceof String) {
            StringBuilder imsgeCss = new StringBuilder(" #");
            imsgeCss.append(name);
            imsgeCss.append("{");
            imsgeCss.append(value);
            imsgeCss.append("}");
            css.insertRule(imsgeCss.toString(), rules.getLength());
        }
        rules = css.getCssRules();
        for (int r = 0; r < rules.getLength(); r++) {
            CSSRule rule = rules.item(r);
            if (rule instanceof CSSStyleRule) {
                CSSStyleRule styleRule = (CSSStyleRule) rule;
                if (styleRule.getSelectorText().equals("*#" + name)) {
                    CSSStyleDeclaration styleDeclaration = styleRule.getStyle();
                    for (int j = 0; j < styleDeclaration.getLength(); j++) {
                        String propertyName = styleDeclaration.item(j);
                        if ("left".equals(propertyName) || "top".equals(propertyName)) {
                            String propertyValue = styleDeclaration
                                    .getPropertyValue(propertyName);
                            Rectangle2D anchor = picture.getAnchor();
                            if ("left".equals(propertyName)) {
                                anchor.setRect(parseLength(propertyValue), anchor.getY(),
                                        anchor.getWidth(), anchor.getHeight());
                            } else {
                                anchor.setRect(anchor.getX(), parseLength(propertyValue),
                                        anchor.getWidth(), anchor.getHeight());
                            }
                            picture.setAnchor(anchor);
                        }
                    }
                }
            }
        }
    }

    public static void embedVideo(XSLFPictureShape picture, Pair<Path, Path> videoPair) throws IOException, InvalidFormatException {
        Path thumbPath = videoPair.getKey();
        byte[] pictureBytes = IOUtils.toByteArray(Files.newInputStream(thumbPath));
        picture.getPictureData().setData(pictureBytes);

        Path videoPath = videoPair.getValue();
        CTPicture xmlObject = (CTPicture) picture.getXmlObject();
        CTApplicationNonVisualDrawingProps nvPr = xmlObject.getNvPicPr().getNvPr();
        String videoId = nvPr.getVideoFile().getLink();

        PackagePart p = picture.getSheet().getPackagePart();
        PackageRelationship rel = p.getRelationship(videoId);

        PackagePart imgPart = p.getRelatedPart(rel);
        XSLFPictureData videoData = new XSLFPictureData(imgPart);
        byte[] videoBytes = IOUtils.toByteArray(Files.newInputStream(videoPath));
        videoData.setData(videoBytes);
    }

    private static double parseLength(String propertyValue) {
        return Double.parseDouble(propertyValue.replace("mm", ""));
    }
}
