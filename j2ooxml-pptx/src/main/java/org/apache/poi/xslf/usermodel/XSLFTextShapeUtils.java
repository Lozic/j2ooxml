package org.apache.poi.xslf.usermodel;

public class XSLFTextShapeUtils {

    public static void copy(XSLFTextShape src, XSLFTextShape dest) {
        if (src.getAnchor() == null) {
            src.setAnchor(dest.getAnchor());
        }
        dest.copy(src);
    }
}
