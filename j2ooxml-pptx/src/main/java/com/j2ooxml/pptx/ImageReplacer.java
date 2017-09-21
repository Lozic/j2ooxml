package com.j2ooxml.pptx;

import java.io.IOException;
import java.nio.file.FileSystem;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.apache.commons.imaging.ImageInfo;
import org.apache.commons.imaging.ImageReadException;
import org.apache.commons.imaging.Imaging;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleRule;
import org.w3c.dom.css.CSSStyleSheet;
import org.xml.sax.SAXException;

import com.j2ooxml.pptx.util.XmlUtil;

/**
 * Replaces image placeholders in pptx by images from model.
 *
 */
public class ImageReplacer {

    private static final String NO_BACKGROUND = "no-background";

    private static final long INCH_TO_EMU = 914400;

    private static final long MM_TO_EMU = 36000;

    public State replace(FileSystem fs, Path slideXml, Map<String, Object> model, CSSStyleSheet css) throws IOException,
            ParserConfigurationException, SAXException, XPathExpressionException, DOMException, ImageReadException {
        Document slideDoc = XmlUtil.parse(slideXml);
        XPath xpath = XPathFactory.newInstance().newXPath();
        XPathExpression expr = xpath.compile("/sld/cSld/spTree/pic/nvPicPr/cNvPr[starts-with(@name, '${') and not(descendant::*)]");
        NodeList nodes = (NodeList) expr.evaluate(slideDoc, XPathConstants.NODESET);

        String slide = slideXml.getFileName().toString();
        Path slideXmlRel = fs.getPath("/ppt/slides/_rels/" + slide + ".rels");
        Document relsDoc = XmlUtil.parse(slideXmlRel);

        processBackground(model, slideDoc);
        processNodes(fs, model, css, xpath, nodes, relsDoc);
        return new State(slideDoc, relsDoc);
    }

    private void processNodes(FileSystem fs, Map<String, Object> model, CSSStyleSheet css, XPath xpath, NodeList nodes, Document relsDoc)
            throws XPathExpressionException, IOException, ImageReadException {
        XPathExpression expr;
        for (int i = 0; i < nodes.getLength(); i++) {
            Element node = (Element) nodes.item(i);
            String name = node.getAttribute("name");
            if (name.endsWith("}")) {
                name = name.substring(2, name.length() - 1);
                Object imageData = model.get(name);
                if (imageData != null) {
                    if (imageData instanceof Path) {
                        Path imagePath = (Path) imageData;
                        String imageId = node.getParentNode().getNextSibling().getFirstChild().getAttributes()
                                .getNamedItem("r:embed").getNodeValue();
                        expr = xpath.compile("//Relationship[@Id='" + imageId + "']");
                        NodeList rels = (NodeList) expr.evaluate(relsDoc, XPathConstants.NODESET);
                        String zipPath = rels.item(0).getAttributes().getNamedItem("Target").getNodeValue();
                        Path fileInsideZipPath = fs.getPath(zipPath.replaceFirst("\\.\\.", "/ppt"));

                        Files.copy(imagePath, fileInsideZipPath, StandardCopyOption.REPLACE_EXISTING);

                    }
                    Element cNvPicPr = (Element) node.getNextSibling();
                    boolean verticalCenter = !cNvPicPr.hasAttribute("preferRelativeResize")
                            || "1".equals(cNvPicPr.getAttribute("preferRelativeResize"));
                    Element picLocks = (Element) cNvPicPr.getFirstChild();
                    boolean smartStretch = !"1".equals(picLocks.getAttribute("noChangeAspect"));

                    Element pic = (Element) node.getParentNode().getParentNode();
                    Element xfrm = (Element) pic.getElementsByTagName("p:spPr").item(0).getFirstChild();
                    Element off = (Element) xfrm.getElementsByTagName("a:off").item(0);

                    if (imageData instanceof Path && (smartStretch || verticalCenter)) {
                        Path imagePath = (Path) imageData;
                        Element ext = (Element) xfrm.getElementsByTagName("a:ext").item(0);

                        long x = Long.parseLong(off.getAttribute("x"));
                        long y = Long.parseLong(off.getAttribute("y"));

                        long wp = Long.parseLong(ext.getAttribute("cx"));
                        long hp = Long.parseLong(ext.getAttribute("cy"));

                        ImageInfo imageInfo = Imaging.getImageInfo(imagePath.toFile());
                        long wi = Math.round(imageInfo.getPhysicalWidthInch() * INCH_TO_EMU);
                        long hi = Math.round(imageInfo.getPhysicalHeightInch() * INCH_TO_EMU);

                        long w = wp;
                        long h = hp;
                        long dx = 0;
                        long dy = 0;
                        if (verticalCenter) {
                            w = wi;
                            h = hi;
                            dx = wp - wi;
                            dy = (hp - hi) / 2;
                        } else if (smartStretch) {
                            if ((float) wp / hp > (float) wi / hi) {
                                w = wi * hp / hi;
                                dx = (wp - w) / 2;
                            } else if ((float) wp / hp < (float) wi / hi) {
                                h = hi * wp / wi;
                                dy = (hp - h) / 2;
                            }
                        }
                        off.setAttribute("x", "" + (x + dx));
                        off.setAttribute("y", "" + (y + dy));
                        ext.setAttribute("cx", "" + w);
                        ext.setAttribute("cy", "" + h);

                    } else {
                        CSSRuleList rules = css.getCssRules();
                        if (imageData instanceof String) {
                            StringBuilder imsgeCss = new StringBuilder(" #");
                            imsgeCss.append(name);
                            imsgeCss.append("{");
                            imsgeCss.append(imageData);
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
                                            String propertyValue = styleDeclaration.getPropertyValue(propertyName);
                                            if ("left".equals(propertyName)) {
                                                off.setAttribute("x", parseLength(propertyValue));
                                            } else {
                                                off.setAttribute("y", parseLength(propertyValue));
                                            }
                                        }
                                    }
                                }
                            }
                        }

                    }

                } else {
                    Node pic = node.getParentNode().getParentNode();
                    Node spTree = pic.getParentNode();
                    spTree.removeChild(pic);
                }
            }
        }
    }

    private void processBackground(Map<String, Object> model, Document slideDoc) {
        Boolean noBackground = (Boolean) model.get(NO_BACKGROUND);
        if (noBackground != null && noBackground) {
            NodeList bgs = slideDoc.getElementsByTagName("p:bg");
            if (bgs.getLength() > 0) {
                Node bg = bgs.item(0);
                bg.getParentNode().removeChild(bg);
            }
        }
    }

    private String parseLength(String propertyValue) {
        Float len = Float.parseFloat(propertyValue.replace("mm", "")) * MM_TO_EMU;
        return "" + len.intValue();
    }
}
