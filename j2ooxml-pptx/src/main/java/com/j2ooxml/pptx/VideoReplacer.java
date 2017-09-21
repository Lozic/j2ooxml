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

import org.apache.commons.imaging.ImageReadException;
import org.apache.commons.lang3.tuple.Pair;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.xml.sax.SAXException;

/**
 * Replaces video placeholders in pptx by video from model.
 *
 */
public class VideoReplacer {

    public void replace(FileSystem fs, Path slideXml, State state, Map<String, Object> model) throws IOException,
            ParserConfigurationException, SAXException, XPathExpressionException, DOMException, ImageReadException {
        Document slideDoc = state.getSlideDoc();
        XPath xpath = XPathFactory.newInstance().newXPath();
        XPathExpression expr = xpath.compile("/sld/cSld/spTree/pic/nvPicPr/cNvPr[starts-with(@name, '${') and (descendant::*)]");
        NodeList nodes = (NodeList) expr.evaluate(slideDoc, XPathConstants.NODESET);
        Document relsDoc = state.getRelDoc();

        processNodes(fs, model, xpath, nodes, slideDoc, relsDoc);
    }

    private void processNodes(FileSystem fs, Map<String, Object> model, XPath xpath, NodeList nodes, Document slideDoc, Document relsDoc)
            throws XPathExpressionException, IOException {
        int length = nodes.getLength();
        int deleted = 0;
        for (int i = 0; i < length; i++) {
            Element node = (Element) nodes.item(i);
            String name = node.getAttribute("name");
            if (name.endsWith("}")) {
                name = name.substring(2, name.length() - 1);
                Object imageData = model.get(name);
                if (imageData != null) {
                    if (imageData instanceof Pair) {
                        @SuppressWarnings("unchecked")
                        Pair<Path, Path> videoPair = (Pair<Path, Path>) imageData;
                        Path thumbPath = videoPair.getKey();
                        Path videoPath = videoPair.getValue();

                        String videoId = node.getNextSibling().getNextSibling().getFirstChild().getAttributes().getNamedItem("r:link")
                                .getNodeValue();
                        replaceMedia(fs, xpath, relsDoc, videoPath, videoId);

                        String thumbId = node.getParentNode().getNextSibling().getFirstChild().getAttributes().getNamedItem("r:embed")
                                .getNodeValue();
                        replaceMedia(fs, xpath, relsDoc, thumbPath, thumbId);
                    }
                } else {
                    Node pic = node.getParentNode().getParentNode();
                    Node spTree = pic.getParentNode();
                    spTree.removeChild(pic);
                    deleted++;
                }
            }
        }
        if (length == deleted) {
            NodeList videoContorlElements = slideDoc.getElementsByTagName("p:timing");
            int videoContorlLength = videoContorlElements.getLength();
            Node sldNode = slideDoc.getDocumentElement();
            for (int i = 0; i < videoContorlLength; i++) {
                Node videoContorlNode = videoContorlElements.item(i);
                sldNode.removeChild(videoContorlNode);
            }
        }
    }

    private void replaceMedia(FileSystem fs, XPath xpath, Document relsDoc, Path path, String id) throws XPathExpressionException, IOException {
        XPathExpression expr = xpath.compile("//Relationship[@Id='" + id + "']");
        NodeList rels = (NodeList) expr.evaluate(relsDoc, XPathConstants.NODESET);
        String zipPath = rels.item(0).getAttributes().getNamedItem("Target").getNodeValue();
        Path fileInsideZipPath = fs.getPath(zipPath.replaceFirst("\\.\\.", "/ppt"));
        Files.copy(path, fileInsideZipPath, StandardCopyOption.REPLACE_EXISTING);
    }
}
