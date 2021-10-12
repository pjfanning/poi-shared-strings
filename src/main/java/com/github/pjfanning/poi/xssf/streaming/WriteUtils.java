package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.util.XMLHelper;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerException;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;
import java.io.IOException;
import java.io.StringReader;
import java.io.StringWriter;

class WriteUtils {

    static String stripXmlFragmentElement(String xml) throws IOException, SAXException, ParserConfigurationException, TransformerException {
        String startTag = "xml-fragment";
        int pos = xml.indexOf(startTag);
        if (pos >= 0) {
            Document doc = XMLHelper.getDocumentBuilderFactory().newDocumentBuilder().parse(new InputSource(new StringReader(xml)));
            NodeList list = doc.getDocumentElement().getChildNodes();
            for (int i = 0; i < list.getLength(); i++) {
                if (list.item(i) instanceof Element) {
                    try (StringWriter sw = new StringWriter()) {
                        Transformer transformer = XMLHelper.newTransformer();
                        transformer.setOutputProperty(OutputKeys.OMIT_XML_DECLARATION, "yes");
                        transformer.transform(new DOMSource(list.item(i)), new StreamResult(sw));
                        return sw.toString();
                    }
                }
            }
            return xml.substring(pos + startTag.length());
        }
        return xml;
    }
}
