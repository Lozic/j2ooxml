package com.j2ooxml.pptx;

import java.lang.reflect.InvocationTargetException;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.apache.commons.beanutils.PropertyUtils;
import org.apache.commons.lang3.StringUtils;
import org.w3c.dom.DOMException;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;
import org.w3c.dom.css.CSSStyleSheet;

import com.j2ooxml.pptx.css.Style;
import com.j2ooxml.pptx.html.Html2PptxTransformer;
import com.j2ooxml.pptx.html.Transformer;

public class VariableProcessor {

    private Transformer transformer = new Html2PptxTransformer();

    public void process(State state, CSSStyleSheet css, Map<String, Object> model) throws GenerationException {
        XPathExpression expr;
        NodeList nodes;
        Pattern varPattern = Pattern.compile("^\\$\\{([a-zA-Z\\._].*?)\\}$");
        XPath xpath1 = XPathFactory.newInstance().newXPath();
        try {
            expr = xpath1.compile("//*[starts-with(text(), '${')]");
            nodes = (NodeList) expr.evaluate(state.getSlideDoc(), XPathConstants.NODESET);
            for (int i = 0; i < nodes.getLength(); i++) {
                String var = null;
                Element node = (Element) nodes.item(i);
                String text = node.getTextContent();

                if ("${".equals(text)) {
                    Node parentNode = node.getParentNode();
                    Node nextParentSibling = parentNode.getNextSibling();
                    if (nextParentSibling != null) {
                        Node nextNextSibling = nextParentSibling.getNextSibling();
                        if (nextNextSibling != null && "}".equals(nextNextSibling.getTextContent())) {
                            var = nextParentSibling.getTextContent();
                            Node parentParent = parentNode.getParentNode();
                            parentParent.removeChild(parentNode);
                            parentParent.removeChild(nextNextSibling);
                            node = (Element) nextParentSibling.getLastChild();
                        }
                    }
                } else {
                    Matcher m = varPattern.matcher(text);
                    if (m.matches()) {
                        var = m.group(1);
                    }
                }
                if (var != null) {
                    evaluate(state, css, model, node, var);
                }
            }
        } catch (XPathExpressionException | DOMException e) {
            throw new GenerationException("Exception while working with pptx inner format.", e);
        }
    }

    private void evaluate(State state, CSSStyleSheet css, Map<String, Object> model, Element node, String var)
            throws GenerationException {
        Object value = null;

        String[] vars = var.split("\\.");
        value = model.get(vars[0]);
        if (value != null) {
            for (int k = 1; k < vars.length; k++) {
                try {
                    value = PropertyUtils.getProperty(value, vars[k]);
                } catch (IllegalAccessException | InvocationTargetException | NoSuchMethodException e) {
                    throw new GenerationException("Could not extract ${'" + var + "} from model.", e);
                }
            }
        }
        try {
            if (StringUtils.isBlank((CharSequence) value)) {
                node.setTextContent(" ");
            } else {

                Node first = node.getParentNode().getParentNode();
                Element rPr = (Element) node.getPreviousSibling();
                String sz = rPr.getAttribute("sz");
                Element txBody = (Element) first.getParentNode();
                state.setTxBody(txBody);
                txBody.removeChild(first);
                String htmlString = (String) value;
                Style style = new Style();
                state.setStyle(style);
                if (StringUtils.isNotBlank(sz)) {
                    try {
                        style.setFontSize(Integer.parseInt(sz));
                    } catch (NumberFormatException e) {
                        throw new GenerationException("Incorrect font size in pptx: " + sz);
                    }
                }
                transformer.convert(state, css, htmlString);
            }
        } catch (DOMException e) {
            throw new GenerationException("Exception while working with pptx inner format.", e);
        }
    }
}
