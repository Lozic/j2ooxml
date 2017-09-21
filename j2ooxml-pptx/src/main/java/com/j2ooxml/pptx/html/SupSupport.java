package com.j2ooxml.pptx.html;

import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;

public class SupSupport implements NodeSupport {

    private Transformer transformer;

    public SupSupport(Transformer transformer) {
        super();
        this.transformer = transformer;
    }

    @Override
    public boolean supports(Node node) {
        if (node instanceof Element) {
            String tagName = ((Element) node).tagName();
            return "sup".equals(tagName);
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        state.getStyle().setBaseline(25000);
        transformer.iterate(state, node);
    }

}
