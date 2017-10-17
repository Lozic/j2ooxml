package com.j2ooxml.pptx.util;

import java.awt.Color;
import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Function;

import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.imaging.ImageInfo;
import org.apache.commons.imaging.ImageReadException;
import org.apache.commons.imaging.Imaging;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.PaintStyle.SolidPaint;
import org.apache.poi.sl.usermodel.Placeholder;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xslf.usermodel.XSLFTextShapeUtils;
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

import com.j2ooxml.pptx.GenerationException;

public class PptxUtil {

    public static void removeSlideBackground(XSLFSlide slide) {
        XSLFBackground bg = slide.getBackground();
        CTBackground xmlBg = (CTBackground) bg.getXmlObject();
        Node bgDomNode = xmlBg.getDomNode();
        bgDomNode.getParentNode().removeChild(bgDomNode);
    }

    public static void embedPicture(XSLFPictureShape picture, Path picturePath) throws IOException, ImageReadException {
        byte[] pictureBytes = IOUtils.toByteArray(Files.newInputStream(picturePath));
        XSLFPictureData pictureData = picture.getPictureData();
        pictureData.setData(pictureBytes);
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
                physicalHeightDpi = CssUtil.DEFAULT_DPI;
            }
            int physicalWidthDpi = imageInfo.getPhysicalWidthDpi();
            if (physicalWidthDpi < 0) {
                physicalWidthDpi = CssUtil.DEFAULT_DPI;
            }
            double wi = Math.round(CssUtil.DEFAULT_DPI * imageInfo.getWidth() / physicalHeightDpi);
            double hi = Math.round(CssUtil.DEFAULT_DPI * imageInfo.getHeight() / physicalWidthDpi);

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

    public static void applyPictureCss(XSLFPictureShape picture, CSSStyleSheet css, String name, Object value) throws GenerationException {
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
                            double length = CssUtil.parseLength(propertyValue);
                            if ("left".equals(propertyName)) {
                                anchor.setRect(length, anchor.getY(),
                                        anchor.getWidth(), anchor.getHeight());
                            } else {
                                anchor.setRect(anchor.getX(), length, anchor.getWidth(), anchor.getHeight());
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

    public static void merge(List<Path> slidesPaths, Path outputPath) throws IOException {
        XMLSlideShow ppt = new XMLSlideShow();
        for (Path slidePath : slidesPaths) {
            InputStream is = Files.newInputStream(slidePath);
            XMLSlideShow src = new XMLSlideShow(is);
            is.close();
            ppt.setPageSize(src.getPageSize());
            for (XSLFSlide srcSlide : src.getSlides()) {
                XSLFSlide destSlide = ppt.createSlide().importContent(srcSlide);
                fixHyperlink(srcSlide, destSlide);
                fixNotes(srcSlide, destSlide);
            }
            Files.delete(slidePath);
            src.close();
        }
        OutputStream out = Files.newOutputStream(outputPath);
        ppt.write(out);
        ppt.close();
        out.close();
    }

    private static void fixNotes(XSLFSlide srcSlide, XSLFSlide destSlide) {
        XSLFNotes srcNotes = srcSlide.getNotes();
        if (srcNotes != null) {
            XSLFNotes destNotes = destSlide.getSlideShow().getNotesSlide(destSlide);
            XSLFTextShape srcNotesBody = getNotesBody(srcNotes);
            XSLFTextShape destNotesBody = getNotesBody(destNotes);
            XSLFTextShapeUtils.copy(srcNotesBody, destNotesBody);
        }
    }

    private static XSLFTextShape getNotesBody(XSLFNotes srcNotes) {
        XSLFTextShape srcNotesBody = null;
        for (XSLFTextShape placeholder : srcNotes.getPlaceholders()) {
            if (Placeholder.BODY.equals(placeholder.getPlaceholder())) {
                srcNotesBody = placeholder;
                break;
            }
        }
        return srcNotesBody;
    }

    /**
     * fix for https://bz.apache.org/bugzilla/show_bug.cgi?id=61589
     */
    private static void fixHyperlink(XSLFSlide srcSlide, XSLFSlide destSlide) {
        List<XSLFHyperlink> srcHyperlinks = getHyperlinks(srcSlide);
        List<XSLFHyperlink> destHyperlinks = getHyperlinks(destSlide);
        if (srcHyperlinks.size() > 0) {
            for (int i = 0; i < srcHyperlinks.size(); i++) {
                XSLFHyperlink srcHyperlink = srcHyperlinks.get(i);
                XSLFHyperlink destHyperlink = destHyperlinks.get(i);
                destHyperlink.linkToUrl(srcHyperlink.getAddress());
            }
        }
    }

    private static List<XSLFHyperlink> getHyperlinks(XSLFSlide slide) {
        List<XSLFHyperlink> hyperlinks = new ArrayList<>();
        for (XSLFShape shape : slide.getShapes()) {
            if (shape instanceof XSLFTextShape) {
                XSLFTextShape textShape = (XSLFTextShape) shape;
                for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                    for (XSLFTextRun run : paragraph.getTextRuns()) {
                        XSLFHyperlink hyperlink = run.getHyperlink();
                        if (hyperlink != null) {
                            hyperlinks.add(hyperlink);
                        }
                    }
                }
            }
        }
        return hyperlinks;
    }

    private static <R> R getDefaultTextShapeValue(XSLFTextShape textShape, Function<XSLFTextParagraph, R> fn) {
        List<XSLFTextParagraph> textParagraphs = textShape.getTextParagraphs();
        R defaultValue = null;
        if (CollectionUtils.isNotEmpty(textParagraphs)) {
            XSLFTextParagraph textParagraph = textParagraphs.get(0);
            defaultValue = fn.apply(textParagraph);
        }
        return defaultValue;
    }

    public static String getDefaultFontFamily(XSLFTextShape textShape) {
        return getDefaultTextShapeValue(textShape, p -> p.getDefaultFontFamily());
    }

    // public static String getDefaultFontFamily(XSLFTextShape textShape) {
    // List<XSLFTextParagraph> textParagraphs = textShape.getTextParagraphs();
    // String defaultFontFamily = null;
    // if (CollectionUtils.isNotEmpty(textParagraphs)) {
    // XSLFTextParagraph textParagraph = textParagraphs.get(0);
    // defaultFontFamily = textParagraph.getDefaultFontFamily();
    // }
    // return defaultFontFamily;
    // }

    public static Double getDefaultFontSize(XSLFTextShape textShape) {
        List<XSLFTextParagraph> textParagraphs = textShape.getTextParagraphs();
        Double defaultFontSize = null;
        if (CollectionUtils.isNotEmpty(textParagraphs)) {
            XSLFTextParagraph textParagraph = textParagraphs.get(0);
            defaultFontSize = textParagraph.getDefaultFontSize();
        }
        return defaultFontSize;
    }

    public static TextAlign getDefaultTextAlign(XSLFTextShape textShape) {
        List<XSLFTextParagraph> textParagraphs = textShape.getTextParagraphs();
        TextAlign defaultTextAlign = null;
        if (CollectionUtils.isNotEmpty(textParagraphs)) {
            XSLFTextParagraph textParagraph = textParagraphs.get(0);
            defaultTextAlign = textParagraph.getTextAlign();
        }
        return defaultTextAlign;
    }

    public static Color getDefaultTextColor(XSLFTextShape textShape) {
        List<XSLFTextParagraph> textParagraphs = textShape.getTextParagraphs();
        Color defaultTextColor = null;
        if (CollectionUtils.isNotEmpty(textParagraphs)) {
            XSLFTextParagraph textParagraph = textParagraphs.get(0);
            PaintStyle fontColor = null;
            List<XSLFTextRun> textRuns = textParagraph.getTextRuns();
            if (textRuns.size() > 0) {
                fontColor = textRuns.get(0).getFontColor();
            }
            if (fontColor instanceof SolidPaint) {
                SolidPaint solidPaint = (SolidPaint) fontColor;
                defaultTextColor = solidPaint.getSolidColor().getColor();
            }
        }
        return defaultTextColor;
    }
}
