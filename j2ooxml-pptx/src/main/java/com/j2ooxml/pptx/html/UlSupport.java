package com.j2ooxml.pptx.html;

import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;
import org.w3c.dom.Document;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;

public class UlSupport implements NodeSupport {

    private Transformer transformer;

    public UlSupport(Transformer transformer) {
        super();
        this.transformer = transformer;
    }

    @Override
    public boolean supports(Node node) {
        if (node instanceof Element) {
            return "ul".equals(((Element) node).tagName());
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        Document slideDoc = state.getSlideDoc();
        org.w3c.dom.Element txBody = state.getTxBody();
        if (node.previousSibling() != null) {
            org.w3c.dom.Element p = slideDoc.createElement("a:p");
            txBody.appendChild(p);
            state.setP(p);
        }
        transformer.iterate(state, node);
    }

}
