package com.j2ooxml.pptx.html;

import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;

public class USupport implements NodeSupport {

    private Transformer transformer;

    public USupport(Transformer transformer) {
        super();
        this.transformer = transformer;
    }

    @Override
    public boolean supports(Node node) {
        if (node instanceof Element) {
            return "u".equals(((Element) node).tagName());
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        state.getStyle().setUnderline(true);
        transformer.iterate(state, node);
    }

}
