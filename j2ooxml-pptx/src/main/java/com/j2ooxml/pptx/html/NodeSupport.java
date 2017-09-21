package com.j2ooxml.pptx.html;

import org.jsoup.nodes.Node;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;

public interface NodeSupport {

    boolean supports(Node node);

    void process(State state, Node node) throws GenerationException;

}
