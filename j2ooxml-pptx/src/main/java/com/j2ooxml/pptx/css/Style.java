package com.j2ooxml.pptx.css;

import java.awt.Color;

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