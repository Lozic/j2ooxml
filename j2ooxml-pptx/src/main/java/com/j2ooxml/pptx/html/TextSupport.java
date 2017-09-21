package com.j2ooxml.pptx.html;

import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;
import org.w3c.dom.DOMException;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;
import com.j2ooxml.pptx.css.Style;
import com.j2ooxml.pptx.css.TextAlign;

public class TextSupport implements NodeSupport {

    @Override
    public boolean supports(Node node) {
        return node instanceof TextNode;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        Element p = state.getP();
        Style style = state.getStyle();
        Document slideDoc = state.getSlideDoc();
        Element r = slideDoc.createElement("a:r");
        p.appendChild(r);
        Element rPr = createRPr(slideDoc, r);

        if (style.getBaseline() != 0) {
            rPr.setAttribute("baseline", "" + style.getBaseline());
        }

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
        TextAlign textAlign = style.getTextAlign();
        if (textAlign != null) {
            Element pPr = (Element) r.getParentNode().getFirstChild();
            textAlign.apply(pPr);
        }
        Element t = slideDoc.createElement("a:t");
        r.appendChild(t);
        t.setTextContent(((TextNode) node).text());
    }

    private static Element createRPr(Document slideDoc, Element e) throws DOMException {
        Element rPr = slideDoc.createElement("a:rPr");
        e.appendChild(rPr);
        rPr.setAttribute("lang", "en-US");
        rPr.setAttribute("dirty", "0");
        rPr.setAttribute("smtClean", "0");
        return rPr;
    }
}
