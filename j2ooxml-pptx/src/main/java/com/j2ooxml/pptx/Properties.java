package com.j2ooxml.pptx;

import java.io.IOException;
import java.nio.file.FileSystem;
import java.nio.file.Path;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.Map;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactoryConfigurationError;

import org.apache.commons.lang3.StringUtils;
import org.w3c.dom.Document;
import org.xml.sax.SAXException;

import com.j2ooxml.pptx.util.XmlUtil;

public final class Properties {

    private static final String PROPERTY_PREFIX = "pptx-";

    public static final String TITLE = PROPERTY_PREFIX + "title";

    public static final String CREATOR = PROPERTY_PREFIX + "creator";

    public static final String LAST_MODIFIED_BY = PROPERTY_PREFIX + "lastModifiedBy";

    public static final String CREATED = PROPERTY_PREFIX + "created";

    public static final String MODIFIED = PROPERTY_PREFIX + "modified";

    public static final String CATEGORY = PROPERTY_PREFIX + "category";

    public static final String KEYWORDS = PROPERTY_PREFIX + "keywords";

    private static final List<String> DC_PROPERTIES = Collections.unmodifiableList(Arrays.asList(TITLE, CREATOR));

    private static final String DC_PREFIX = "dc:";

    private static final List<String> CP_PROPERTIES = Collections
            .unmodifiableList(Arrays.asList(LAST_MODIFIED_BY, KEYWORDS, CATEGORY));

    private static final String CP_PREFIX = "cp:";

    private static final List<String> DCTERMS_PROPERTIES = Collections
            .unmodifiableList(Arrays.asList(CREATED, MODIFIED));

    private static final String DCTERMS_PREFIX = "dcterms:";

    public static void fillProperies(FileSystem fs, Map<String, Object> model)
            throws IOException, ParserConfigurationException, SAXException,
            TransformerConfigurationException, TransformerFactoryConfigurationError, TransformerException {
        Path coreXml = fs.getPath("/docProps/core.xml");
        Document coreDoc = XmlUtil.parse(coreXml);
        fillProperties(model, coreDoc, DC_PROPERTIES, DC_PREFIX);
        fillProperties(model, coreDoc, CP_PROPERTIES, CP_PREFIX);
        fillProperties(model, coreDoc, DCTERMS_PROPERTIES, DCTERMS_PREFIX);
        XmlUtil.save(coreXml, coreDoc);
    }

    public static void fillProperties(Map<String, Object> model, Document document, List<String> properties,
            String prefix) {
        properties.forEach(prop -> {
            if (model.containsKey(prop)) {
                String value = model.get(prop).toString();
                String property = prop.replace(PROPERTY_PREFIX, StringUtils.EMPTY);
                document.getElementsByTagName(prefix + property).item(0).setTextContent(value);
            }
        });
    }

    private Properties() {
    }
}
