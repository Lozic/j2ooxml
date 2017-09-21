package com.j2ooxml.pptx;

import org.w3c.dom.Document;
import org.w3c.dom.Element;

import com.j2ooxml.pptx.css.Style;

/**
 * Represents {@link PptxGenerator} inner state.
 *
 */
public class State {

    private Document slideDoc;
    private Document relDoc;
    private Element txBody;
    private Element p;
    private Style style;

    public State(Document slideDoc, Document relDoc) {
        this.slideDoc = slideDoc;
        this.relDoc = relDoc;
    }

    public Document getSlideDoc() {
        return slideDoc;
    }

    public void setSlideDoc(Document slideDoc) {
        this.slideDoc = slideDoc;
    }

    public Document getRelDoc() {
        return relDoc;
    }

    public void setRelDoc(Document relDoc) {
        this.relDoc = relDoc;
    }

    public Element getTxBody() {
        return txBody;
    }

    public void setTxBody(Element txBody) {
        this.txBody = txBody;
    }

    public Element getP() {
        return p;
    }

    public void setP(Element p) {
        this.p = p;
    }

    public Style getStyle() {
        return style;
    }

    public void setStyle(Style style) {
        this.style = style;
    }

}