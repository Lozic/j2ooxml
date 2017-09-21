package com.j2ooxml.pptx;

import java.util.Map;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

public class EmptyNodesRemover {

    public void process(Document slideDoc, Map<String, Object> model) throws GenerationException {
        XPathExpression expr;
        NodeList nodes;
        XPath xpath = XPathFactory.newInstance().newXPath();
        try {
            if (model.containsKey("customer_title") && model.get("customer_title") == null) {
                expr = xpath.compile("/sldLayout/cSld/spTree/sp/txBody/p/r/t[text()='Customers']/../../../..");
                nodes = (NodeList) expr.evaluate(slideDoc, XPathConstants.NODESET);
                for (int i = 0; i < nodes.getLength(); i++) {
                    Element node = (Element) nodes.item(i);
                    node.getParentNode().removeChild(node);
                }
            }
            if (model.containsKey("kf_title") && model.get("kf_title") == null) {
                expr = xpath.compile("/sldLayout/cSld/spTree/sp/txBody/p/r/t[text()='Key facts']/../../../..");
                nodes = (NodeList) expr.evaluate(slideDoc, XPathConstants.NODESET);
                for (int i = 0; i < nodes.getLength(); i++) {
                    Element node = (Element) nodes.item(i);
                    node.getParentNode().removeChild(node);
                }
            }
        } catch (XPathExpressionException | DOMException e) {
            throw new GenerationException("Exception while working with pptx inner format.", e);
        }
    }

}
