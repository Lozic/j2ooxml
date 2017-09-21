package com.j2ooxml.pptx.css;

import java.awt.Color;
import java.lang.reflect.Field;

import org.apache.commons.lang3.StringUtils;

import com.j2ooxml.pptx.GenerationException;

public class CssProcessor {

    public static void process(String htmlStyle, Style style) throws GenerationException {
        try {

            if (StringUtils.isNoneBlank(htmlStyle)) {
                for (String css : htmlStyle.split(";")) {
                    String[] cssVal = css.split(":");
                    String value = cssVal[1].trim();
                    switch (cssVal[0]) {
                    case "color":
                        value = extractColor(value);
                        style.setColor(value);
                        break;
                    case "font-size":
                        style.setFontSize(Math.round(100 * Float.parseFloat(value.replace("pt", "").trim())));
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
                        style.setLiColor(extractColor(value));
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
                            style.setTextAlign(TextAlign.of(value));
                        }
                        break;
                    default:
                        break;
                    }
                }
            }
        } catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
            throw new GenerationException("Incorrect css style for element.", e);
        }
    }

    private static String extractColor(String color)
            throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException {
        if (color.startsWith("#")) {
            color = color.replace("#", "").toUpperCase();
        } else if (color.startsWith("rgb")) {
            String[] rgb = color.replace("rgb", "").replace("(", "").replace(")", "").split(",");
            color = String.format("%02X%02X%02X", Integer.parseInt(rgb[0].trim()), Integer.parseInt(rgb[1].trim()),
                    Integer.parseInt(rgb[2].trim()));
        } else {
            Field f = Color.class.getField(color);
            Color col = (Color) f.get(null);
            color = String.format("%02X%02X%02X", col.getRed(), col.getGreen(), col.getBlue());
        }
        return color;
    }
}
