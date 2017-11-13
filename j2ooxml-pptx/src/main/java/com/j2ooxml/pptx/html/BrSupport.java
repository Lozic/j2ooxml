package com.j2ooxml.pptx.html;

import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.nodes.Element;
import org.jsoup.nodes.Node;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;

public class BrSupport implements NodeSupport {

    @Override
    public boolean supports(Node node) {
        if (node instanceof Element) {
            return "br".equals(((Element) node).tagName());
        }
        return false;
    }

    @Override
    public void process(State state, Node node) throws GenerationException {
        XSLFTextShape textShape = state.getTextShape();
        XSLFTextParagraph paragraph = textShape.addNewTextParagraph();
        paragraph.setSpaceAfter(0.);
        paragraph.setSpaceBefore(0.);
        state.setParagraph(paragraph);
    }

}
