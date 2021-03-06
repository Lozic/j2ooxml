package com.j2ooxml.pptx.html;

import java.lang.reflect.InvocationTargetException;
import java.util.HashSet;
import java.util.Set;
import java.util.stream.Collectors;
import java.util.stream.Stream;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.jsoup.Jsoup;
import org.w3c.dom.css.CSSStyleSheet;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.State;
import com.j2ooxml.pptx.css.Style;
import com.j2ooxml.pptx.util.CssUtil;

public class Html2PptxTransformer implements Transformer {

    private Set<NodeSupport> supportSet;

    public Html2PptxTransformer() {
        super();
        supportSet = Stream.of(
                new ASupport(),
                new BrSupport(),
                new BSupport(this),
                new ISupport(this),
                new LiSupport(this),
                new SubSupport(this),
                new SupSupport(this),
                new TextSupport(),
                new UlSupport(this),
                new USupport(this))
                .collect(Collectors.toCollection(HashSet::new));
    }

    @Override
    public void convert(State state, CSSStyleSheet css, String htmlString) throws GenerationException {
        org.jsoup.nodes.Document html = Jsoup.parse(htmlString);
        CssUtil.applyCss(css, html);
        org.jsoup.nodes.Node body = html.body();

        XSLFTextShape shape = state.getTextShape();
        XSLFTextParagraph paragraph = shape.addNewTextParagraph();
        state.setParagraph(paragraph);

        Style style = state.getStyle();
        TextAlign textAlign = style.getTextAlign();
        if (textAlign != null) {
            paragraph.setTextAlign(textAlign);
        }
        iterate(state, body);

    }

    @Override
    public void iterate(State state, org.jsoup.nodes.Node node) throws GenerationException {
        try {
            for (org.jsoup.nodes.Node htmlNode : node.childNodes()) {
                traverse(state, htmlNode);
            }
        } catch (Exception e) {
            throw new GenerationException(e);
        }
    }

    private void traverse(State state, org.jsoup.nodes.Node node) throws GenerationException {
        try {
            Style style = state.getStyle();
            Style prevStyle = (Style) BeanUtils.cloneBean(style);
            String htmlStyle = node.attr("style");
            CssUtil.process(htmlStyle, style);

            NodeSupport support = supports(node);

            if (support != null) {
                support.process(state, node);
            } else {
                iterate(state, node);
            }

            state.setStyle(prevStyle);
        } catch (GenerationException e) {
            throw e;
        } catch (IllegalAccessException | InstantiationException | InvocationTargetException | NoSuchMethodException e) {
            throw new GenerationException(e);
        }
    }

    private NodeSupport supports(org.jsoup.nodes.Node node) {
        NodeSupport support = null;
        for (NodeSupport sup : supportSet) {
            if (sup.supports(node)) {
                support = sup;
                break;
            }
        }
        return support;
    }

}
