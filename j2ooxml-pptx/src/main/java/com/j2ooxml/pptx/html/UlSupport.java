package com.j2ooxml.pptx.html;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;

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
        XSLFTextParagraph paragraph = state.getParagraph();
        if (node.previousSibling() != null) {
            paragraph = state.getTextShape().addNewTextParagraph();
            paragraph.setSpaceAfter(0.);
            paragraph.setSpaceBefore(0.);
            state.setParagraph(paragraph);
        }
        transformer.iterate(state, node);
    }

}
