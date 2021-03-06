package com.j2ooxml.pptx;

import java.awt.Color;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.imaging.ImageReadException;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JRuntimeException;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFNotes;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTable;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.presentationml.x2006.main.CTApplicationNonVisualDrawingProps;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;
import org.w3c.dom.css.CSSStyleSheet;

import com.j2ooxml.pptx.css.Style;
import com.j2ooxml.pptx.html.Html2PptxTransformer;
import com.j2ooxml.pptx.html.Transformer;
import com.j2ooxml.pptx.util.CssUtil;
import com.j2ooxml.pptx.util.PptxPropertiesUtils;
import com.j2ooxml.pptx.util.PptxUtil;

public class PptxTemplate {

    private static final String NO_BACKGROUND = "no-background";

    public static void generate(Path templatePath, Path cssPath, Path outputPath, Map<String, Object> model) throws IOException, GenerationException {
        try {
            Files.copy(templatePath, outputPath, StandardCopyOption.REPLACE_EXISTING);
            CSSStyleSheet css = CssUtil.parseCss(cssPath, model);
            XMLSlideShow ppt = new XMLSlideShow(Files.newInputStream(outputPath));
            PptxPropertiesUtils.fillProperies(ppt, model);
            for (XSLFSlide slide : ppt.getSlides()) {
                processSlide(slide, model, css);
            }
            OutputStream out = Files.newOutputStream(outputPath);
            ppt.write(out);
            out.close();
            ppt.close();
        } catch (IOException | ImageReadException | InvalidFormatException e) {
            throw new GenerationException("Could not generate resulting ppt.", e);
        }
    }

    private static void processSlide(XSLFSlide slide, Map<String, Object> model, CSSStyleSheet css)
            throws GenerationException, IOException, ImageReadException, InvalidFormatException {
        Boolean noBackground = (Boolean) model.get(NO_BACKGROUND);
        if (noBackground != null && noBackground) {
            PptxUtil.removeSlideBackground(slide);
        }
        Set<XSLFShape> shpesToRemove = new HashSet<>();
        int videoCount = 0;
        for (XSLFShape sh : slide.getShapes()) {
            videoCount += processShape(sh, model, css, shpesToRemove);
        }
        for (XSLFShape sh : shpesToRemove) {
            try {
                slide.removeShape(sh);
            } catch (OpenXML4JRuntimeException e) {
                slide.removeShape(sh);
            }
        }
        if (videoCount <= 0) {
            CTSlide xslide = slide.getXmlObject();
            if (xslide.isSetTiming()) {
                xslide.unsetTiming();
            }
        }
        processNotes(slide, model, css);
    }

    private static void processNotes(XSLFSlide slide, Map<String, Object> model, CSSStyleSheet css) throws GenerationException {
        XSLFNotes notes = slide.getNotes();
        if (notes != null) {
            for (XSLFShape shape : notes.getShapes()) {
                if (shape instanceof XSLFTextShape) {
                    XSLFTextShape textShape = (XSLFTextShape) shape;
                    String text = textShape.getText();
                    text = text.trim();
                    if (StringUtils.isNotEmpty(text) && text.startsWith("${") && text.endsWith("}")) {
                        text = text.substring(2, text.length() - 1);
                        String value = (String) model.get(text);
                        if (value == null) {
                            value = StringUtils.EMPTY;
                        }
                        processTextShape(textShape, value, css);
                    }
                }
            }
        }
    }

    private static int processShape(XSLFShape sh, Map<String, Object> model, CSSStyleSheet css, Set<XSLFShape> shpesToRemove)
            throws GenerationException, IOException, ImageReadException, InvalidFormatException {
        int videoCount = 0;
        String name = sh.getShapeName();
        if (StringUtils.isNotEmpty(name) && name.startsWith("${") && name.endsWith("}")) {
            name = name.substring(2, name.length() - 1);
            Object value = null;
            String[] vars = name.split("\\.");
            boolean present = model.containsKey(vars[0]);
            value = model.get(vars[0]);
            int len = vars.length;
            if (value != null && len > 1) {
                for (int k = 1; k < len; k++) {
                    try {
                        value = PropertyUtils.getProperty(value, vars[k]);
                        present = true;
                    } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                        throw new GenerationException("Could not extract ${'" + name + "} from model.", e);
                    }
                }
            }
            if (present && value == null) {
                shpesToRemove.add(sh);
            }
            if (sh instanceof XSLFTextShape || sh instanceof XSLFTable) {
                XSLFTextShape textShape = getTextShape(sh);
                processTextShape(textShape, value, css);
            } else if (sh instanceof XSLFPictureShape) {
                if (!present || value == null) {
                    shpesToRemove.add(sh);
                } else {
                    XSLFPictureShape picture = (XSLFPictureShape) sh;
                    CTPicture xmlObject = (CTPicture) picture.getXmlObject();
                    CTApplicationNonVisualDrawingProps nvPr = xmlObject.getNvPicPr().getNvPr();
                    boolean video = nvPr.isSetVideoFile();
                    if (video && shpesToRemove.contains(sh)) {
                        videoCount--;
                    }
                    if (value instanceof Path) {
                        Path picturePath = (Path) value;
                        PptxUtil.embedPicture(picture, picturePath);
                    } else if (value instanceof Pair<?, ?>) {
                        if (video) {
                            @SuppressWarnings("unchecked")
                            Pair<Path, Path> videoPair = (Pair<Path, Path>) value;
                            PptxUtil.embedVideo(picture, videoPair);
                            videoCount++;
                        } else {
                            @SuppressWarnings("unchecked")
                            Pair<Path, String> imgCssPair = (Pair<Path, String>) value;
                            PptxUtil.embedPicture(picture, imgCssPair.getLeft());
                            PptxUtil.applyPictureCss(picture, css, name, imgCssPair.getRight());
                        }
                    } else {
                        PptxUtil.applyPictureCss(picture, css, name, value);
                    }
                }
            }
        }
        return videoCount;
    }

    private static void processTextShape(XSLFTextShape textShape, Object value, CSSStyleSheet css) throws GenerationException {
        Transformer transformer = new Html2PptxTransformer();
        String defaultFontFamily = PptxUtil.getDefaultFontFamily(textShape);
        Double defaultFontSize = PptxUtil.getDefaultFontSize(textShape);
        TextAlign defaultTextAlign = PptxUtil.getDefaultTextAlign(textShape);
        Color defaultColor = PptxUtil.getDefaultTextColor(textShape);
        textShape.clearText();
        if (!StringUtils.isBlank((CharSequence) value)) {
            String htmlString = (String) value;
            Style style = new Style();
            State state = new State(textShape);
            state.setStyle(style);
            if (defaultFontFamily != null) {
                style.setFontFamily(defaultFontFamily);
            }
            if (defaultFontSize != null) {
                style.setFontSize(defaultFontSize);
            }
            if (defaultTextAlign != null) {
                style.setTextAlign(defaultTextAlign);
            }
            if (defaultColor != null) {
                style.setColor(defaultColor);
            }
            transformer.convert(state, css, htmlString);
        } else {
            textShape.setText(" ");
        }
    }

    private static XSLFTextShape getTextShape(XSLFShape sh) {
        XSLFTextShape textShape;
        if (sh instanceof XSLFTextShape) {
            textShape = (XSLFTextShape) sh;
        } else {
            XSLFTable table = (XSLFTable) sh;
            textShape = table.getCell(0, 0);
        }
        return textShape;
    }
}
