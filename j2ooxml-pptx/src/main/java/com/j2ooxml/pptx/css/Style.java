package com.j2ooxml.pptx.css;

import java.awt.Color;
import java.lang.reflect.Field;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;

/**
 * Simple aggregator for css style.
 *
 */
public class Style {

    private boolean bold;
    private boolean italic;
    private boolean underline;
    private Color color;
    private Double fontSize;
    private Color liColor;
    private String liChar;
    private double indent;
    private double marginLeft;
    private TextAlign textAlign;
    private int baseline;

    public boolean isBold() {
        return bold;
    }

    public void setBold(boolean bold) {
        this.bold = bold;
    }

    public boolean isItalic() {
        return italic;
    }

    public void setItalic(boolean italic) {
        this.italic = italic;
    }

    public boolean isUnderline() {
        return underline;
    }

    public void setUnderline(boolean underline) {
        this.underline = underline;
    }

    public Color getColor() {
        return color;
    }

    public void setColor(Color color) {
        this.color = color;
    }

    public void setColor(String color) {
        this.color = parseColor(color);
    }

    public Double getFontSize() {
        return fontSize;
    }

    public void setFontSize(Double fontSize) {
        this.fontSize = fontSize;
    }

    public Color getLiColor() {
        return liColor;
    }

    public void setLiColor(Color liColor) {
        this.liColor = liColor;
    }

    public void setLiColor(String color) {
        this.liColor = parseColor(color);
    }

    private Color parseColor(String color) {
        Color result = null;
        color = color.trim();
        if (StringUtils.isNoneEmpty(color)) {
            color = color.replace(" ", "").toLowerCase();
            int length = color.length();
            if (color.startsWith("rgb")) {
                if (color.startsWith("rgb(")) {
                    color = color.substring(4, length);
                    String[] rgb = color.split(",");
                    result = new Color(Integer.parseInt(rgb[0]), Integer.parseInt(rgb[1]), Integer.parseInt(rgb[2]));
                } else {
                    color = color.substring(5, length);
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

    public String getLiChar() {
        return liChar;
    }

    public void setLiChar(String liChar) {
        this.liChar = liChar;
    }

    public double getIndent() {
        return indent;
    }

    public void setIndent(double indent) {
        this.indent = indent;
    }

    public double getMarginLeft() {
        return marginLeft;
    }

    public void setMarginLeft(double marginLeft) {
        this.marginLeft = marginLeft;
    }

    public TextAlign getTextAlign() {
        return textAlign;
    }

    public void setTextAlign(TextAlign textAlign) {
        this.textAlign = textAlign;
    }

    public long getBaseline() {
        return baseline;
    }

    public void setBaseline(int baseline) {
        this.baseline = baseline;
    }
}