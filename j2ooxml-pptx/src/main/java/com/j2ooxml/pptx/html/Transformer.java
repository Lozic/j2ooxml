package com.j2ooxml.pptx.html;

import org.jsoup.nodes.Node;
import org.w3c.dom.css.CSSStyleSheet;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;

public interface Transformer {

    void iterate(State state, Node node) throws GenerationException;

    void convert(State state, CSSStyleSheet css, String htmlString) throws GenerationException;

}
