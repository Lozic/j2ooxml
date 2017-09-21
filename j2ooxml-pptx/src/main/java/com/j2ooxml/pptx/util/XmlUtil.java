package com.j2ooxml.pptx.util;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.Result;
import javax.xml.transform.Source;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerConfigurationException;
import javax.xml.transform.TransformerException;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.TransformerFactoryConfigurationError;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.w3c.dom.Document;
import org.xml.sax.SAXException;

public final class XmlUtil {

    public static void save(Path path, Document document) throws TransformerConfigurationException,
            TransformerFactoryConfigurationError, IOException, TransformerException {
        Transformer transformer = TransformerFactory.newInstance().newTransformer();
        Path temp = Files.createTempFile(null, ".xml");
        Result output = new StreamResult(Files.newOutputStream(temp));
        Source input = new DOMSource(document);
        transformer.transform(input, output);
        Files.copy(temp, path, StandardCopyOption.REPLACE_EXISTING);
        Files.delete(temp);
    }

    public static Document parse(Path path) throws IOException, ParserConfigurationException, SAXException {
        InputStream xmlInput = Files.newInputStream(path);
        DocumentBuilderFactory factory = DocumentBuilderFactory.newInstance();
        DocumentBuilder builder = factory.newDocumentBuilder();
        Document coreDoc = builder.parse(xmlInput);
        return coreDoc;
    }

    private XmlUtil() {
    }

}
