package com.j2ooxml.pptx;

import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringReader;
import java.lang.reflect.InvocationTargetException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.imaging.ImageInfo;
import org.apache.commons.imaging.Imaging;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.tuple.Pair;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.openxml4j.opc.PackageRelationship;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualPictureProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTApplicationNonVisualDrawingProps;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackground;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPictureNonVisual;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;
import org.w3c.css.sac.InputSource;
import org.w3c.dom.DOMException;
import org.w3c.dom.Node;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleRule;
import org.w3c.dom.css.CSSStyleSheet;

import com.j2ooxml.pptx.css.Style;
import com.j2ooxml.pptx.html.Html2PptxTransformer;
import com.j2ooxml.pptx.html.Transformer;
import com.steadystate.css.parser.CSSOMParser;

public class PptxGenerator {

    private static final String NO_BACKGROUND = "no-background";

    public void process(Path templatePath, Path cssPath, Path outputPath, Map<String, Object> model)
            throws IOException, GenerationException {
        try {
            Files.copy(templatePath, outputPath, StandardCopyOption.REPLACE_EXISTING);
            CSSOMParser parser = new CSSOMParser();
            String cssString = new String(Files.readAllBytes(cssPath), StandardCharsets.UTF_8);
            if (model.containsKey("CSS")) {
                String modelCss = (String) model.get("CSS");
                cssString += " " + modelCss;
            }
            StringReader reader = new StringReader(cssString);
            CSSStyleSheet css = parser.parseStyleSheet(new InputSource(reader), null, null);

            XMLSlideShow ppt = new XMLSlideShow(Files.newInputStream(outputPath));

            PptxProperties.fillProperies(ppt, model);

            for (XSLFSlide slide : ppt.getSlides()) {
                Boolean noBackground = (Boolean) model.get(NO_BACKGROUND);
                if (noBackground) {
                    XSLFBackground bg = slide.getBackground();
                    CTBackground xmlBg = (CTBackground) bg.getXmlObject();
                    Node bgDomNode = xmlBg.getDomNode();
                    bgDomNode.getParentNode().removeChild(bgDomNode);
                }

                Set<XSLFShape> shpesToRemove = new HashSet<>();
                int videoCount = 0;
                for (XSLFShape sh : slide.getShapes()) {
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
                                } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                                    throw new GenerationException("Could not extract ${'" + name + "} from model.", e);
                                }
                            }
                        }

                        if (present && value == null) {
                            shpesToRemove.add(sh);
                        }
                        if (sh instanceof XSLFTextShape) {
                            XSLFTextShape textShape = (XSLFTextShape) sh;

                            Transformer transformer = new Html2PptxTransformer();

                            try {
                                List<XSLFTextParagraph> textParagraphs = textShape.getTextParagraphs();
                                Double defaultFontSize = null;
                                if (CollectionUtils.isNotEmpty(textParagraphs)) {
                                    XSLFTextParagraph textParagraph = textParagraphs.get(0);
                                    defaultFontSize = textParagraph.getDefaultFontSize();
                                }
                                textShape.clearText();
                                if (!StringUtils.isBlank((CharSequence) value)) {
                                    String htmlString = (String) value;
                                    Style style = new Style();
                                    State state = new State(textShape);
                                    state.setStyle(style);
                                    if (defaultFontSize != null) {
                                        style.setFontSize(defaultFontSize);
                                    }
                                    transformer.convert(state, css, htmlString);
                                } else {
                                    textShape.setText(" ");
                                }
                            } catch (DOMException e) {
                                throw new GenerationException("Exception while working with pptx inner format.", e);
                            }

                        } else if (sh instanceof XSLFPictureShape) {
                            if (!present || value == null) {
                                shpesToRemove.add(sh);
                            } else {
                                XSLFPictureShape picture = (XSLFPictureShape) sh;
                                Rectangle2D anchor = picture.getAnchor();
                                CTPicture xmlObject = (CTPicture) picture.getXmlObject();
                                CTPictureNonVisual nvPicPr = xmlObject.getNvPicPr();
                                CTNonVisualPictureProperties cNvPicPr = nvPicPr.getCNvPicPr();
                                boolean verticalCenter = !cNvPicPr.isSetPreferRelativeResize() || cNvPicPr.getPreferRelativeResize();
                                boolean smartStretch = !cNvPicPr.getPicLocks().getNoChangeAspect();

                                CTApplicationNonVisualDrawingProps nvPr = xmlObject.getNvPicPr().getNvPr();
                                boolean video = nvPr.isSetVideoFile();
                                if (video && shpesToRemove.contains(sh)) {
                                    videoCount--;
                                }

                                if (value instanceof Path) {
                                    Path picturePath = (Path) value;
                                    byte[] pictureBytes = IOUtils.toByteArray(Files.newInputStream(picturePath));
                                    picture.getPictureData().setData(pictureBytes);
                                }
                                if (value instanceof Path && (smartStretch || verticalCenter)) {
                                    Path picturePath = (Path) value;
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
                                } else if (value instanceof Pair<?, ?>) {
                                    @SuppressWarnings("unchecked")
                                    Pair<Path, Path> videoPair = (Pair<Path, Path>) value;

                                    Path thumbPath = videoPair.getKey();
                                    byte[] pictureBytes = IOUtils.toByteArray(Files.newInputStream(thumbPath));
                                    picture.getPictureData().setData(pictureBytes);

                                    Path videoPath = videoPair.getValue();
                                    String videoId = nvPr.getVideoFile().getLink();

                                    PackagePart p = sh.getSheet().getPackagePart();
                                    PackageRelationship rel = p.getRelationship(videoId);

                                    PackagePart imgPart = p.getRelatedPart(rel);
                                    XSLFPictureData videoData = new XSLFPictureData(imgPart);
                                    byte[] videoBytes = IOUtils.toByteArray(Files.newInputStream(videoPath));
                                    videoData.setData(videoBytes);
                                    videoCount++;

                                } else {
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
                                                        if ("left".equals(propertyName)) {
                                                            anchor.setRect(parseLength(propertyValue), anchor.getY(),
                                                                    anchor.getWidth(), anchor.getHeight());
                                                        } else {
                                                            anchor.setRect(anchor.getX(), parseLength(propertyValue),
                                                                    anchor.getWidth(), anchor.getHeight());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                }
                            }

                        }
                    }
                }
                for (XSLFShape sh : shpesToRemove) {
                    slide.removeShape(sh);
                }
                if (videoCount <= 0) {
                    CTSlide xslide = slide.getXmlObject();
                    xslide.unsetTiming();
                }
            }

            OutputStream out = Files.newOutputStream(outputPath);
            ppt.write(out);
            out.close();
            ppt.close();

        } catch (Exception e) {
            throw new GenerationException("Could not generate resulting ppt.", e);
        }
    }

    private double parseLength(String propertyValue) {
        return Double.parseDouble(propertyValue.replace("mm", ""));
    }
}
