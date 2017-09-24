package com.j2ooxml.pptx;

import static j2html.TagCreator.body;
import static j2html.TagCreator.div;
import static j2html.TagCreator.head;
import static j2html.TagCreator.html;
import static j2html.TagCreator.img;
import static j2html.TagCreator.p;
import static j2html.TagCreator.span;
import static j2html.TagCreator.style;

import java.awt.Color;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.util.List;
import java.util.concurrent.ThreadLocalRandom;

import org.apache.poi.sl.usermodel.PlaceableShape;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFConnectorShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSimpleShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import j2html.tags.ContainerTag;
import j2html.tags.EmptyTag;

public class PptxTest {
	public static void main(String[] args) throws FileNotFoundException, IOException {

        XMLSlideShow ppt = new XMLSlideShow(new FileInputStream("btsf2-1.pptx"));
        java.awt.Dimension pgsize = ppt.getPageSize();
        int pgx = pgsize.width;
        int pgy = pgsize.height;
        String slideStyle = "width: " + dotsToCm(pgx) + "; height: " + dotsToCm(pgy) + ";";

        ContainerTag html = html();
        ContainerTag style = style();
        style.withText(".pptx-slide {\r\n" +
                "    position: relative;\r\n" +
                "}\r\n" +
                "\r\n" +
                ".pptx-slide div {\r\n" +
                "    position: absolute;\r\n" +
                "}\r\n" +
                "");
        html.with(head().with(style));
        ContainerTag body = body();
        html.with(body);
        // get slides
        for (XSLFSlide slide : ppt.getSlides()) {
            Color fillColor = slide.getBackground().getFillColor();
            String backgroud = colorToRgba(fillColor);
            ContainerTag slideDiv = div().withStyle(slideStyle + backgroud).withClass("pptx-slide");
            for (XSLFSlideMaster masterSlide : ppt.getSlideMasters()) {
                parseSlide(slideDiv, masterSlide.getShapes());
                for (XSLFSlideLayout slideLayout : masterSlide.getSlideLayouts()) {
                    parseSlide(slideDiv, slideLayout.getShapes());
                }
            }
            parseSlide(slideDiv, slide.getShapes());
            body.with(slideDiv);
        }
        String formatted = html.renderFormatted();
        System.out.println(formatted);
        Files.write(Paths.get("test.html"), formatted.getBytes(), StandardOpenOption.CREATE, StandardOpenOption.TRUNCATE_EXISTING);
        ppt.close();
    }

    private static String colorToRgba(Color color) {
        int red = color.getRed();
        int green = color.getGreen();
        int blue = color.getBlue();
        int alpha = color.getAlpha();
        StringBuilder rgba = new StringBuilder("background-color: ");
        rgba.append(alpha == 255 ? "rgb(" : "rgba(");
        rgba.append(red);
        rgba.append(",");
        rgba.append(green);
        rgba.append(",");
        rgba.append(blue);
        if (alpha < 255) {
            rgba.append(",");
            rgba.append(alpha / 255.);
        }
        rgba.append(");");
        String backgroud = rgba.toString();
        return backgroud;
    }

    private static ContainerTag parseSlide(ContainerTag slideDiv, List<XSLFShape> shapes) throws IOException {
        ThreadLocalRandom rnd = ThreadLocalRandom.current();
        for (XSLFShape sh : shapes) {
            if (sh instanceof PlaceableShape) {
                java.awt.geom.Rectangle2D anchor = ((PlaceableShape<?, ?>) sh).getAnchor();
                String shapeStyle = "left: " + dotsToCm(anchor.getX()) + "; top: " + dotsToCm(anchor.getY())
                        + "; width: " + dotsToCm(anchor.getWidth()) + "; height: " + dotsToCm(anchor.getHeight()) + ";";
                // TODO: uncoment for random background
                // + "; background-color: rgb(" + rnd.nextInt(256) + "," + rnd.nextInt(256) + "," + rnd.nextInt(256) + ");";
                if (sh instanceof XSLFSimpleShape) {
                    XSLFSimpleShape shape = (XSLFSimpleShape) sh;
                    Color fillColor = shape.getFillColor();
                    if (fillColor != null) {
                        shapeStyle += colorToRgba(fillColor);
                    }
                }
                ContainerTag div = div().withStyle(shapeStyle);
                slideDiv.with(div);
                // TODO: grouping shapes children positions does not work correctly in current POI
                // if (sh instanceof XSLFGroupShape) {
                // XSLFGroupShape groupShape = (XSLFGroupShape) sh;
                // parseSlide(div, groupShape.getShapes());
                // }
                if (sh instanceof XSLFTextShape) {
                    XSLFTextShape shape = (XSLFTextShape) sh;
                    for (XSLFTextParagraph paragraph : shape.getTextParagraphs()) {
                        String defaultFontFamily = paragraph.getDefaultFontFamily();
                        Double defaultFontSize = paragraph.getDefaultFontSize();
                        String pStyle = "margin-top: 0 ;font-family: " + defaultFontFamily + "; font-size: " + defaultFontSize + "pt";
                        ContainerTag p = p().withStyle(pStyle);
                        div.with(p);
                        for (XSLFTextRun run : paragraph.getTextRuns()) {
                            String fontFamily = run.getFontFamily();
                            Double fontSize = run.getFontSize();
                            String spanStyle = "font-family: " + fontFamily + "; font-size: " + fontSize + "pt;";
                            if (run.isBold()) {
                                spanStyle = spanStyle + "font-weight: bold;";
                            }
                            ContainerTag span = span().withStyle(spanStyle).withText(run.getRawText());
                            p.with(span);
                        }
                    }
                } else if (sh instanceof XSLFPictureShape) {
                    XSLFPictureShape pictureShape = (XSLFPictureShape) sh;
                    XSLFPictureData pictureData = pictureShape.getPictureData();
                    String fileName = pictureData.getFileName();
                    InputStream pictureStream = pictureData.getInputStream();
                    Path imagePath = Paths.get(fileName);
                    Files.copy(pictureStream, imagePath, StandardCopyOption.REPLACE_EXISTING);
                    String imgStyle = "width: 100%;  height: 100%;";
                    EmptyTag img = img().withSrc(fileName).withStyle(imgStyle);
                    div.with(img);
                }
            }

            if (sh instanceof XSLFConnectorShape) {
                XSLFConnectorShape line = (XSLFConnectorShape) sh;
                // work with Line
            }
        }
        return slideDiv;
    }

    private static String dotsToCm(double px) {
        double cm = 2.54 * px / 72;
        BigDecimal bcm = new BigDecimal(cm);
        bcm = bcm.setScale(2, RoundingMode.HALF_UP);
        return bcm.doubleValue() + "cm";
    }
}
