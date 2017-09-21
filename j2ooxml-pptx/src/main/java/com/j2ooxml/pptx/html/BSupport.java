package com.j2ooxml.pptx.html;

import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;

public class BSupport implements NodeSupport {

    private Transformer transformer;

    public BSupport(Transformer transformer) {
        super();
        this.transformer = transformer;
    }

    @Override
    public boolean supports(Node node) {
        if (node instanceof Element) {
            String tagName = ((Element) node).tagName();
            return "b".equals(tagName) || "strong".equals(tagName);
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        state.getStyle().setBold(true);
        transformer.iterate(state, node);
    }

}
