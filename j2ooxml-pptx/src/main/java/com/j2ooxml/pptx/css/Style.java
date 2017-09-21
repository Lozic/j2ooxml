package com.j2ooxml.pptx.css;

/**
 * Simple aggregator for css style.
 *
 */
public class Style {

    private boolean bold;
    private boolean italic;
    private boolean underline;
    private String color;
    private int fontSize;
    private String liColor;
    private String liChar;
    private float indent;
    private float marginLeft;
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

    public String getColor() {
        return color;
    }

    public void setColor(String color) {
        this.color = color;
    }

    public int getFontSize() {
        return fontSize;
    }

    public void setFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public String getLiColor() {
        return liColor;
    }

    public void setLiColor(String liColor) {
        this.liColor = liColor;
    }

    public String getLiChar() {
        return liChar;
    }

    public void setLiChar(String liChar) {
        this.liChar = liChar;
    }

    public float getIndent() {
        return indent;
    }

    public void setIndent(float indent) {
        this.indent = indent;
    }

    public float getMarginLeft() {
        return marginLeft;
    }

    public void setMarginLeft(float marginLeft) {
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