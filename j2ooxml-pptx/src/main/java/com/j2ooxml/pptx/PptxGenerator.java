package com.j2ooxml.pptx;

import java.awt.geom.Rectangle2D;
import java.io.IOException;
import java.io.OutputStream;
import java.io.StringReader;
import java.nio.charset.StandardCharsets;
import java.nio.file.DirectoryStream;
import java.nio.file.FileSystem;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.commons.imaging.ImageInfo;
import org.apache.commons.imaging.Imaging;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.sl.usermodel.PlaceableShape;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFConnectorShape;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualPictureProperties;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPicture;
import org.openxmlformats.schemas.presentationml.x2006.main.CTPictureNonVisual;
import org.w3c.css.sac.InputSource;
import org.w3c.dom.Document;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleRule;
import org.w3c.dom.css.CSSStyleSheet;

import com.j2ooxml.pptx.util.XmlUtil;
import com.steadystate.css.parser.CSSOMParser;

public class PptxGenerator {

    private VariableProcessor variableProcessor = new VariableProcessor();

    private ImageReplacer imageReplacer = new ImageReplacer();

    private VideoReplacer videoReplacer = new VideoReplacer();

    private Deletor deletor = new Deletor();

    private EmptyNodesRemover emptyNodesRemover = new EmptyNodesRemover();

    public void process(Path templatePath, Path cssPath, Path outputPath, Map<String, Object> model)
            throws IOException, GenerationException {
        try {
            Files.copy(templatePath, outputPath, StandardCopyOption.REPLACE_EXISTING);
            FileSystem fs = FileSystems.newFileSystem(outputPath, null);

            CSSOMParser parser = new CSSOMParser();
            String cssString = new String(Files.readAllBytes(cssPath), StandardCharsets.UTF_8);
            if (model.containsKey("CSS")) {
                String modelCss = (String) model.get("CSS");
                cssString += " " + modelCss;
            }
            StringReader reader = new StringReader(cssString);
            CSSStyleSheet css = parser.parseStyleSheet(new InputSource(reader), null, null);

            Path slides = fs.getPath("/ppt/slides");
            try (DirectoryStream<Path> directoryStream = Files.newDirectoryStream(slides)) {
                for (Path slideXml : directoryStream) {
                    if (Files.isRegularFile(slideXml)) {
                        String slide = slideXml.getFileName().toString();
                        Path relXml = fs.getPath("/ppt/slides/_rels/" + slide + ".rels");

                        State state = imageReplacer.replace(fs, slideXml, model, css);
                        videoReplacer.replace(fs, slideXml, state, model);
                        deletor.process(state, model);
                        variableProcessor.process(state, css, model);
                        XmlUtil.save(slideXml, state.getSlideDoc());
                        XmlUtil.save(relXml, state.getRelDoc());
                        try (Stream<String> lines = Files.lines(relXml, StandardCharsets.UTF_8)) {
                            List<String> replaced = lines
                                    .map(line -> line.replaceAll(
                                            "(Target=\".*?\"(\\sTargetMode=\".*?\")?)\\s(Type=\".*?\")", "$3 $1"))
                                    .collect(Collectors.toList());
                            lines.close();
                            Files.delete(relXml);
                            Files.write(relXml, replaced, StandardCharsets.UTF_8);
                        }
                    }
                }
            }

            slides = fs.getPath("/ppt/slideLayouts");
            try (DirectoryStream<Path> directoryStream = Files.newDirectoryStream(slides)) {
                for (Path slideXml : directoryStream) {
                    if (Files.isRegularFile(slideXml)) {
                        Document slideDoc = XmlUtil.parse(slideXml);
                        emptyNodesRemover.process(slideDoc, model);
                        XmlUtil.save(slideXml, slideDoc);
                    }
                }
            }
            Properties.fillProperies(fs, model);
            fs.close();

            XMLSlideShow ppt = new XMLSlideShow(Files.newInputStream(outputPath));

            for (XSLFSlide slide : ppt.getSlides()) {
                List<XSLFShape> shpesToRemove = new ArrayList<>();
                for (XSLFShape sh : slide.getShapes()) {
                    String name = sh.getShapeName();
                    if (StringUtils.isNotEmpty(name) && name.startsWith("${") && name.endsWith("}")) {
                        name = name.substring(2, name.length() - 1);
                        Object value = model.get(name);
                        // shapes's anchor which defines the position of this shape in the slide
                        if (sh instanceof PlaceableShape) {
                            java.awt.geom.Rectangle2D anchor = ((PlaceableShape) sh).getAnchor();
                        }

                        if (sh instanceof XSLFConnectorShape) {
                            XSLFConnectorShape line = (XSLFConnectorShape) sh;
                            // work with Line
                        } else if (sh instanceof XSLFTextShape) {
                            XSLFTextShape shape = (XSLFTextShape) sh;
                            // work with a shape that can hold text
                        } else if (sh instanceof XSLFPictureShape) {
                            if (value == null) {
                                shpesToRemove.add(sh);
                            } else {
                                XSLFPictureShape picture = (XSLFPictureShape) sh;
                                Rectangle2D anchor = picture.getAnchor();
                                CTPicture xmlObject = (CTPicture) picture.getXmlObject();
                                CTPictureNonVisual nvPicPr = xmlObject.getNvPicPr();
                                CTNonVisualPictureProperties cNvPicPr = nvPicPr.getCNvPicPr();
                                boolean verticalCenter = !cNvPicPr.isSetPreferRelativeResize() || cNvPicPr.getPreferRelativeResize();
                                boolean smartStretch = !cNvPicPr.getPicLocks().getNoChangeAspect();
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
