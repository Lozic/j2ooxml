package com.j2ooxml.pptx.css;

import java.util.HashMap;
import java.util.Map;

import org.w3c.dom.Element;

public enum TextAlign {
    CENTER("center", "ctr"),
    LEFT("left", "l"),
    RIGHT("right", "r");

    private String css;
    private String pptx;

    private static final Map<String, TextAlign> map;
    static {
        map = new HashMap<>();
        for (TextAlign v : TextAlign.values()) {
            map.put(v.css, v);
        }
    }

    private TextAlign(String css, String pptx) {
        this.css = css;
        this.pptx = pptx;
    }

    public String getCss() {
        return css;
    }

    public String getPptx() {
        return pptx;
    }

    public static TextAlign of(String css) {
        return map.get(css);
    }

    public void apply(Element element) {
        element.setAttribute("algn", pptx);
    }
}
