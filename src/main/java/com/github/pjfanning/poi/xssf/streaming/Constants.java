package com.github.pjfanning.poi.xssf.streaming;

import org.apache.xmlbeans.XmlOptions;

import javax.xml.namespace.QName;
import java.security.SecureRandom;
import java.util.Collections;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

class Constants {
    static final SecureRandom RANDOM = new SecureRandom();

    static final XmlOptions saveOptions = new XmlOptions();
    static {
        saveOptions.setCharacterEncoding("UTF-8");
        saveOptions.setSaveAggressiveNamespaces();
        saveOptions.setUseDefaultNamespace(true);
        saveOptions.setSaveImplicitNamespaces(Collections.singletonMap("", NS_SPREADSHEETML));
    }
}
