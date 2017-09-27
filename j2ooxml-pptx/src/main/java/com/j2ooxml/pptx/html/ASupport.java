package com.j2ooxml.pptx.html;

import org.apache.poi.xslf.usermodel.XSLFHyperlink;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.jsoup.nodes.Node;
import org.jsoup.nodes.TextNode;

public class ASupport extends TextSupport {

    @Override
    public boolean supports(Node node) {
        if (node instanceof org.jsoup.nodes.Element) {
            return "a".equals(((org.jsoup.nodes.Element) node).tagName());
        }
        return false;
    }

    @Override
    protected void setContent(XSLFTextRun textRun, Node node) {
        org.jsoup.nodes.Element link = (org.jsoup.nodes.Element) node;
        String linkText;
        if (link.childNodeSize() == 1 && link.childNode(0) instanceof TextNode) {
            linkText = ((TextNode) link.childNode(0)).getWholeText();
        } else {
            linkText = link.text();
        }
        textRun.setText(linkText);
        XSLFHyperlink hyperlink = textRun.createHyperlink();
        hyperlink.setAddress(link.attr("href"));
    }

}
