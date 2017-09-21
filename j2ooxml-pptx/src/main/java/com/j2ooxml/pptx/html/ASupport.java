package com.j2ooxml.pptx.html;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

import javax.xml.xpath.XPath;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathExpression;
import javax.xml.xpath.XPathExpressionException;
import javax.xml.xpath.XPathFactory;

import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;
import com.j2ooxml.pptx.css.Style;

public class ASupport implements NodeSupport {

    @Override
    public boolean supports(Node node) {
        if (node instanceof org.jsoup.nodes.Element) {
            return "a".equals(((org.jsoup.nodes.Element) node).tagName());
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        Element p = state.getP();
        Document slideDoc = state.getSlideDoc();
        Style style = state.getStyle();
        Element r = slideDoc.createElement("a:r");
        p.appendChild(r);
        Element rPr = createRPr(slideDoc, r);
        if (style.isBold()) {
            rPr.setAttribute("b", "1");
        }
        if (style.isItalic()) {
            rPr.setAttribute("i", "1");
        }
        if (style.isUnderline()) {
            rPr.setAttribute("u", "sng");
        }
        if (style.getColor() != null) {
            Element solidFill = slideDoc.createElement("a:solidFill");
            rPr.appendChild(solidFill);
            Element srgbClr = slideDoc.createElement("a:srgbClr");
            solidFill.appendChild(srgbClr);
            srgbClr.setAttribute("val", style.getColor());
        }
        if (style.getFontSize() > 0) {
            rPr.setAttribute("sz", "" + style.getFontSize());
        }
        Element t = slideDoc.createElement("a:t");
        r.appendChild(t);
        org.jsoup.nodes.Element link = (org.jsoup.nodes.Element) node;

        String linkText;
        if (link.childNodeSize() == 1 && link.childNode(0) instanceof TextNode) {
            linkText = ((TextNode) link.childNode(0)).getWholeText();
        } else {
            linkText = link.text();
        }
        t.setTextContent(linkText);

        try {
            XPath xpath = XPathFactory.newInstance().newXPath();
            XPathExpression expr = xpath.compile("//Relationship");
            Document relDoc = state.getRelDoc();
            NodeList rels = (NodeList) expr.evaluate(relDoc, XPathConstants.NODESET);
            int maxId = 1;
            List<Integer> ids = new ArrayList<>();
            for (int i = 0; i < rels.getLength(); i++) {
                Element rel = (Element) rels.item(i);
                ids.add(Integer.parseInt(rel.getAttribute("Id").substring(3)));
            }
            if (ids.size() > 0) {
                maxId = Collections.max(ids);
            }
            Element relation = relDoc.createElement("Relationship");
            String rId = "rId" + (maxId + 1);
            relation.setAttribute("Id", rId);
            relation.setAttribute("Type",
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink");
            relation.setAttribute("Target", link.attr("href"));
            relation.setAttribute("TargetMode", "External");
            relDoc.getFirstChild().appendChild(relation);

            Element hlinkClick = slideDoc.createElement("a:hlinkClick");
            rPr.appendChild(hlinkClick);
            hlinkClick.setAttribute("r:id", rId);

        } catch (XPathExpressionException e) {
            throw new GenerationException("Error prcessing link tag: " + link.outerHtml(), e);
        }

    }

    private static Element createRPr(Document slideDoc, Element e) throws DOMException {
        Element rPr = slideDoc.createElement("a:rPr");
        e.appendChild(rPr);
        rPr.setAttribute("lang", "en-US");
        rPr.setAttribute("dirty", "0");
        rPr.setAttribute("smtClean", "0");
        // rPr.setAttribute("err", "1");
        return rPr;
    }
}
