package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.util.XMLHelper;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLEventWriter;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;
import java.io.IOException;
import java.io.StringWriter;
import java.util.Collections;

class TextParser {

    static String getXMLText(XMLEventReader xmlEventReader, QName tag) throws IOException, XMLStreamException {
        XMLEventFactory xef = XMLHelper.newXMLEventFactory();
        try (StringWriter sw = new StringWriter()) {
            XMLEventWriter xew = XMLHelper.newXMLOutputFactory().createXMLEventWriter(sw);
            xew.add(xef.createStartElement(tag, Collections.emptyIterator(), Collections.emptyIterator()));
            try {
                XMLEvent event = xmlEventReader.nextEvent();
                while (event != null && !(event.isEndElement() && event.asEndElement().getName().equals(tag))) {
                    xew.add(event);
                    event = xmlEventReader.nextEvent();
                }
                xew.add(xef.createEndElement(tag, Collections.emptyIterator()));
            } finally {
                xew.close();
            }
            return sw.toString();
        }
    }
}
