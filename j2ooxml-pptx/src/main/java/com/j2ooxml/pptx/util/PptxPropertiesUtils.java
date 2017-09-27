package com.j2ooxml.pptx.util;

import java.util.Map;

import org.apache.poi.POIXMLProperties;
import org.apache.poi.POIXMLProperties.CoreProperties;
import org.apache.poi.POIXMLProperties.CustomProperties;
import org.apache.poi.xslf.usermodel.XMLSlideShow;

public final class PptxPropertiesUtils {

    private static final String PROPERTY_PREFIX = "pptx-";
    private static final int PROPERTY_PREFIX_LEN = PROPERTY_PREFIX.length();
    public static final String TITLE = PROPERTY_PREFIX + "title";
    public static final String CREATOR = PROPERTY_PREFIX + "creator";
    public static final String LAST_MODIFIED_BY = PROPERTY_PREFIX + "lastModifiedBy";
    public static final String CREATED = PROPERTY_PREFIX + "created";
    public static final String MODIFIED = PROPERTY_PREFIX + "modified";
    public static final String CATEGORY = PROPERTY_PREFIX + "category";
    public static final String KEYWORDS = PROPERTY_PREFIX + "keywords";

    private PptxPropertiesUtils() {
    }

    public static void addCustomProperty(String propertyName, String propertyValue, Map<String, Object> model) {
        model.put(PROPERTY_PREFIX + propertyName, propertyValue);
    }

    public static void fillProperies(XMLSlideShow ppt, Map<String, Object> model) {
        model.entrySet().stream().filter(e -> e.getKey().startsWith(PROPERTY_PREFIX)).forEach(e -> {
            String key = e.getKey();
            String value = e.getValue().toString();
            POIXMLProperties properties = ppt.getProperties();
            CoreProperties coreProperties = properties.getCoreProperties();
            CustomProperties customProperties = properties.getCustomProperties();
            switch (key) {
            case TITLE:
                coreProperties.setTitle(value);
                break;
            case CREATOR:
                coreProperties.setCreator(value);
                break;
            case LAST_MODIFIED_BY:
                coreProperties.setLastModifiedByUser(value);
                break;
            case CREATED:
                coreProperties.setCreated(value);
                break;
            case MODIFIED:
                coreProperties.setModified(value);
                break;
            case CATEGORY:
                coreProperties.setCategory(value);
                break;
            case KEYWORDS:
                coreProperties.setKeywords(value);
                break;
            default:
                String property = key.substring(PROPERTY_PREFIX_LEN);
                customProperties.addProperty(property, value);
                break;
            }
        });
    }

}
