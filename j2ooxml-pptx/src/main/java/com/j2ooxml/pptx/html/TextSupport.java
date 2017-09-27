package com.j2ooxml.pptx.html;

import java.awt.Color;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;
import com.j2ooxml.pptx.css.Style;

public class TextSupport implements NodeSupport {

    @Override
    public boolean supports(Node node) {
        return node instanceof TextNode;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        XSLFTextParagraph paragraph = state.getParagraph();
        Style style = state.getStyle();
        XSLFTextRun textRun = paragraph.addNewTextRun();
        textRun.setBold(style.isBold());
        textRun.setItalic(style.isItalic());
        textRun.setUnderlined(style.isUnderline());
        Color color = style.getColor();
        if (color != null) {
            textRun.setFontColor(color);
        }
        Double fontSize = style.getFontSize();
        if (fontSize != null) {
            textRun.setFontSize(fontSize);
        }
        textRun.setBaselineOffset(style.getBaseline());
        setContent(textRun, node);
    }

    protected void setContent(XSLFTextRun textRun, Node node) {
        textRun.setText(((TextNode) node).text());
    }
}
