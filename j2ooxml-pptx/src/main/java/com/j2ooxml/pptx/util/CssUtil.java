package com.j2ooxml.pptx.util;

import java.awt.Color;
import java.io.IOException;
import java.io.StringReader;
import java.lang.reflect.Field;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;
import org.w3c.css.sac.InputSource;
import org.w3c.dom.css.CSSRule;
import org.w3c.dom.css.CSSRuleList;
import org.w3c.dom.css.CSSStyleDeclaration;
import org.w3c.dom.css.CSSStyleRule;
import org.w3c.dom.css.CSSStyleSheet;

import com.j2ooxml.pptx.GenerationException;
import com.j2ooxml.pptx.css.Style;
import com.steadystate.css.parser.CSSOMParser;

public class CssUtil {

    public static void applyCss(CSSStyleSheet css, org.jsoup.nodes.Document html) {
        Map<String, Map<String, String>> styles = traverse(css, html);
        apply(styles, html);
    }

    public static void process(String htmlStyle, Style style) throws GenerationException {
        try {

            if (StringUtils.isNoneBlank(htmlStyle)) {
                for (String css : htmlStyle.split(";")) {
                    String[] cssVal = css.split(":");
                    String value = cssVal[1].trim();
                    switch (cssVal[0]) {
                    case "color":
                        style.setColor(parseColor(value));
                        break;
                    case "font-size":
                        style.setFontSize(Double.parseDouble(value.replace("pt", "").trim()));
                        break;
                    case "font-weight":
                        if ("bold".equals(value)) {
                            style.setBold(true);
                        }
                        break;
                    case "li-content":
                        style.setLiChar(String.valueOf(value.charAt(1)));
                        break;
                    case "li-color":
                        style.setLiColor(parseColor(value));
                        break;
                    case "indent":
                        style.setIndent(Float.parseFloat(value.replace("mm", "")));
                        break;
                    case "margin-left":
                        style.setMarginLeft(Float.parseFloat(value.replace("mm", "")));
                        break;
                    case "text-decoration":
                        if ("underline".equals(value)) {
                            style.setUnderline(true);
                        }
                        break;
                    case "text-align":
                        if (StringUtils.isNotBlank(value)) {
                            style.setTextAlign(TextAlign.valueOf(value.toUpperCase()));
                        }
                        break;
                    default:
                        break;
                    }
                }
            }
        } catch (IllegalArgumentException e) {
            throw new GenerationException("Incorrect css style for element.", e);
        }
    }

    public static CSSStyleSheet parseCss(Path cssPath, Map<String, Object> model) throws IOException {
        CSSOMParser parser = new CSSOMParser();
        String cssString = new String(Files.readAllBytes(cssPath), StandardCharsets.UTF_8);
        if (model.containsKey("CSS")) {
            String modelCss = (String) model.get("CSS");
            cssString += " " + modelCss;
        }
        StringReader reader = new StringReader(cssString);
        CSSStyleSheet css = parser.parseStyleSheet(new InputSource(reader), null, null);
        return css;
    }

    private static Color parseColor(String color) {
        Color result = null;
        color = color.trim();
        if (StringUtils.isNoneEmpty(color)) {
            color = color.replace(" ", "").toLowerCase();
            int length = color.length();
            if (color.startsWith("#")) {
                String r = color.substring(1, 3);
                String g = color.substring(3, 5);
                String b = color.substring(5, 7);
                result = new Color(Integer.parseInt(r, 16), Integer.parseInt(g, 16), Integer.parseInt(b, 16));
            } else if (color.startsWith("rgb")) {
                if (color.startsWith("rgb(")) {
                    color = color.substring(4, length - 1);
                    String[] rgb = color.split(",");
                    result = new Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2]));
                } else {
                    color = color.substring(5, length - 1);
                    String[] rgb = color.split(",");
                    int alpha = 255 * (int) Math.round(Double.parseDouble(rgb[3]));
                    result = new Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2]), alpha);
                }
            } else {
                try {
                    final Field f = Color.class.getField(color);
                    result = (Color) f.get(null);
                } catch (Exception ce) {
                    result = null;
                }
            }
        }
        return result;
    }

    private static void apply(Map<String, Map<String, String>> styles, org.jsoup.nodes.Document html) {
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

    private static Map<String, Map<String, String>> traverse(CSSStyleSheet css, org.jsoup.nodes.Document html) {
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

    private static void traverseSelected(org.jsoup.nodes.Document html, Map<String, Map<String, String>> styles,
            CSSStyleRule styleRule, String selector) {
        Elements selectedElements = html.select(selector);
        for (Element selected : selectedElements) {
            traverseElement(styles, styleRule, selected.cssSelector());
        }
    }

    private static void traverseElement(Map<String, Map<String, String>> styles, CSSStyleRule styleRule, String selector) {
        if (!styles.containsKey(selector)) {
            styles.put(selector, new LinkedHashMap<String, String>());
        }
        CSSStyleDeclaration styleDeclaration = styleRule.getStyle();
        for (int j = 0; j < styleDeclaration.getLength(); j++) {
            String propertyName = styleDeclaration.item(j);
            String propertyValue = styleDeclaration.getPropertyValue(propertyName);
            Map<String, String> elementStyle = styles.get(selector);
            elementStyle.put(propertyName, propertyValue);
        }
    }
}
