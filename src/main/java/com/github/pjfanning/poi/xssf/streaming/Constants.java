package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.util.XMLHelper;
import org.apache.xmlbeans.XmlOptions;

import javax.xml.stream.XMLEventFactory;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLOutputFactory;
import java.security.SecureRandom;
import java.util.Collections;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

class Constants {
    static final SecureRandom RANDOM = new SecureRandom();
    static final int DEFAULT_CAFFEINE_CACHE_SIZE = 100;

    static final XmlOptions saveOptions = new XmlOptions();
    static {
        saveOptions.setCharacterEncoding("UTF-8");
        saveOptions.setSaveAggressiveNamespaces();
        saveOptions.setUseDefaultNamespace(true);
        saveOptions.setSaveImplicitNamespaces(Collections.singletonMap("", NS_SPREADSHEETML));
    }

    static final XMLInputFactory XML_INPUT_FACTORY = XMLHelper.newXMLInputFactory();
    static final XMLOutputFactory XML_OUTPUT_FACTORY = XMLHelper.newXMLOutputFactory();
    static final XMLEventFactory XML_EVENT_FACTORY = XMLHelper.newXMLEventFactory();
}
