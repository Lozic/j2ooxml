package com.j2ooxml.pptx;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

import com.j2ooxml.pptx.css.Style;

public class State {

    XSLFTextShape textShape;
    XSLFTextParagraph paragraph;
    private Style style;

    public State(XSLFTextShape textShape) {
        super();
        this.textShape = textShape;
    }

    public XSLFTextShape getTextShape() {
        return textShape;
    }

    public void setTextShape(XSLFTextShape textShape) {
        this.textShape = textShape;
    }

    public XSLFTextParagraph getParagraph() {
        return paragraph;
    }

    public void setParagraph(XSLFTextParagraph paragraph) {
        this.paragraph = paragraph;
    }

    public Style getStyle() {
        return style;
    }

    public void setStyle(Style style) {
        this.style = style;
    }

}