package com.j2ooxml.pptx.html;

import java.awt.Color;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.jsoup.nodes.Node;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;
import com.j2ooxml.pptx.css.Style;

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
        XSLFTextParagraph paragraph = state.getParagraph();
        Style style = state.getStyle();
        TextAlign textAlign = style.getTextAlign();
        if (textAlign != null) {
            paragraph.setTextAlign(textAlign);
        }
        double indent = style.getIndent();
        if (indent > 0) {
            paragraph.setIndent(indent);
        }
        double marginLeft = style.getMarginLeft();
        if (marginLeft > 0) {
            paragraph.setLeftMargin(marginLeft);
            if (indent == 0) {
                paragraph.setIndent(-marginLeft);
            }
        }
        Color liColor = style.getLiColor();
        if (liColor != null) {
            paragraph.setBulletFontColor(liColor);
        }
        String bulletChar = StringUtils.isNotBlank(style.getLiChar()) ? style.getLiChar() : "\u2022";
        paragraph.setBulletCharacter(bulletChar);

        transformer.iterate(state, node);

        paragraph = state.getTextShape().addNewTextParagraph();
        state.setParagraph(paragraph);
        if (node.nextSibling() == null) {
            if (textAlign != null) {
                paragraph.setTextAlign(textAlign);
            }
        }
    }

}
