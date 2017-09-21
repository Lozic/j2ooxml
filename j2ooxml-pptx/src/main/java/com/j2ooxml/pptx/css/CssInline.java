package com.j2ooxml.pptx.css;

import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleRule;
import org.w3c.dom.css.CSSStyleSheet;

public class CssInline {

    public void applyCss(CSSStyleSheet css, org.jsoup.nodes.Document html) {
        Map<String, Map<String, String>> styles = traverse(css, html);
        apply(styles, html);
    }

    private void apply(Map<String, Map<String, String>> styles, org.jsoup.nodes.Document html) {
        for (Entry<String, Map<String, String>> style : styles.entrySet()) {
            String selector = style.getKey();
            Map<String, String> map = style.getValue();
            StringBuilder builder = new StringBuilder();
            for (Entry<String, String> css : map.entrySet()) {
                builder.append(css.getKey()).append(":").append(css.getValue()).append(";");
            }
            Elements elements = html.select(selector);
            for (Element element : elements) {
                builder.append(element.attr("style"));
                element.attr("style", builder.toString());
                element.removeAttr("class");
            }
        }
    }

    private Map<String, Map<String, String>> traverse(CSSStyleSheet css, org.jsoup.nodes.Document html) {
        CSSRuleList rules = css.getCssRules();
        Map<String, Map<String, String>> styles = new HashMap<>();

        for (int i = 0; i < rules.getLength(); i++) {
            CSSRule rule = rules.item(i);
            if (rule instanceof CSSStyleRule) {
                CSSStyleRule styleRule = (CSSStyleRule) rule;
                String selector = styleRule.getSelectorText();

                if (!selector.contains(":")) {
                    traverseSelected(html, styles, styleRule, selector);
                }
            }
        }
        return styles;
    }

    private void traverseSelected(org.jsoup.nodes.Document html, Map<String, Map<String, String>> styles,
            CSSStyleRule styleRule, String selector) {
        final Elements selectedElements = html.select(selector);
        for (Element selected : selectedElements) {
            traverseElement(styles, styleRule, selected.cssSelector());

        }
    }

    private void traverseElement(Map<String, Map<String, String>> styles, CSSStyleRule styleRule, String selector) {
        if (!styles.containsKey(selector)) {
            styles.put(selector, new LinkedHashMap<String, String>());
        }

        final CSSStyleDeclaration styleDeclaration = styleRule.getStyle();

        for (int j = 0; j < styleDeclaration.getLength(); j++) {
            final String propertyName = styleDeclaration.item(j);
            final String propertyValue = styleDeclaration.getPropertyValue(propertyName);
            final Map<String, String> elementStyle = styles.get(selector);
            elementStyle.put(propertyName, propertyValue);
        }
    }
}
