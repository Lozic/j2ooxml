package com.j2ooxml.pptx.html;

import org.apache.commons.lang3.StringUtils;
import org.jsoup.nodes.Node;
import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;
import com.j2ooxml.pptx.css.Style;
import com.j2ooxml.pptx.css.TextAlign;

public class LiSupport implements NodeSupport {

    private Transformer transformer;

    public LiSupport(Transformer transformer) {
        super();
        this.transformer = transformer;
    }

    @Override
    public boolean supports(Node node) {
        if (node instanceof org.jsoup.nodes.Element) {
            return "li".equals(((org.jsoup.nodes.Element) node).tagName());
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        Element p = state.getP();
        Style style = state.getStyle();
        Document slideDoc = state.getSlideDoc();
        Element txBody = state.getTxBody();
        if (node.parent().previousSibling() == null && node.previousSibling() == null) {
            p.removeChild(p.getFirstChild());
        }
        Element pPr = slideDoc.createElement("a:pPr");
        p.appendChild(pPr);
        TextAlign textAlign = style.getTextAlign();
        if (textAlign != null) {
            textAlign.apply(pPr);
        }

        float indent = style.getIndent();
        if (indent > 0) {
            pPr.setAttribute("indent", "" + Math.round(indent * 36000));
        }
        float marginLeft = style.getMarginLeft();
        if (marginLeft > 0) {
            pPr.setAttribute("marL", "" + Math.round(marginLeft * 36000));
            if (indent == 0) {
                pPr.setAttribute("indent", "" + Math.round(-marginLeft * 36000));
            }
        }
        Element lnSpc = slideDoc.createElement("a:lnSpc");
        pPr.appendChild(lnSpc);
        Element spcPct = slideDoc.createElement("a:spcPct");
        lnSpc.appendChild(spcPct);
        spcPct.setAttribute("val", "100000");
        Element spcBef = slideDoc.createElement("a:spcBef");
        pPr.appendChild(spcBef);
        Element spcPts = slideDoc.createElement("a:spcPts");
        spcBef.appendChild(spcPts);
        spcPts.setAttribute("val", "0");

        if (StringUtils.isNotBlank(style.getLiColor())) {
            Element buClr = slideDoc.createElement("a:buClr");
            pPr.appendChild(buClr);
            Element srgbClr = slideDoc.createElement("a:srgbClr");
            buClr.appendChild(srgbClr);
            srgbClr.setAttribute("val", style.getLiColor());
        }
        Element buSzPct = slideDoc.createElement("a:buSzPct");
        pPr.appendChild(buSzPct);
        buSzPct.setAttribute("val", "120000");

        Element buFont = slideDoc.createElement("a:buFont");
        pPr.appendChild(buFont);
        buFont.setAttribute("typeface", "Arial");
        buFont.setAttribute("pitchFamily", "34");
        buFont.setAttribute("charset", "0");

        Element buChar = slideDoc.createElement("a:buChar");
        pPr.appendChild(buChar);
        String bulletChar = StringUtils.isNotBlank(style.getLiChar()) ? style.getLiChar() : "\u2022";
        buChar.setAttribute("char", bulletChar);

        transformer.iterate(state, node);
        p = slideDoc.createElement("a:p");
        txBody.appendChild(p);
        state.setP(p);
        if (node.nextSibling() == null) {
            pPr = slideDoc.createElement("a:pPr");
            p.appendChild(pPr);
            if (textAlign != null) {
                textAlign.apply(pPr);
            }
        }
    }

}
