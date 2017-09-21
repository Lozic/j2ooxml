package com.j2ooxml.pptx;

import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.w3c.dom.DOMException;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

public class Deletor {

    public void process(State state, Map<String, Object> model) throws GenerationException {
        XPathExpression expr;
        NodeList nodes;
        Pattern varPattern = Pattern.compile("^\\$\\{([a-zA-Z\\._].*?)\\}$");
        XPath xpath = XPathFactory.newInstance().newXPath();
        try {
            expr = xpath.compile("/sld/cSld/spTree/*/*/cNvPr[starts-with(@name, '${')]");
            nodes = (NodeList) expr.evaluate(state.getSlideDoc(), XPathConstants.NODESET);
            for (int i = 0; i < nodes.getLength(); i++) {
                String var = null;
                Element node = (Element) nodes.item(i);
                String text = node.getAttribute("name");
                Matcher m = varPattern.matcher(text);
                if (m.matches()) {
                    var = m.group(1);
                }
                if (var != null && model.containsKey(var) && model.get(var) == null) {
                    Node grandPa = node.getParentNode().getParentNode();
                    Node grandGrandPa = grandPa.getParentNode();
                    grandGrandPa.removeChild(grandPa);
                }
            }
        } catch (XPathExpressionException | DOMException e) {
            throw new GenerationException("Exception while working with pptx inner format.", e);
        }
    }

}
