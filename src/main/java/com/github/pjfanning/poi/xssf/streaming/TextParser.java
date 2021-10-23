package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.util.XMLHelper;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventFactory;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLEventWriter;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.EndElement;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import java.io.IOException;
import java.io.StringWriter;
import java.util.Collections;
import java.util.List;
import java.util.ListIterator;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

class TextParser {

    private static final XMLEventFactory xef = XMLHelper.newXMLEventFactory();

    /**
     * Parses a {@code <si>} String Item. Returns just the text and drops the formatting. See <a
     * href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sharedstringitem.aspx">xmlschema
     * type {@code CT_Rst}</a>.
     */
    static String parseCT_Rst(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to <si> or <text>;  Post condition: pointing to </si> or </text>
        StringBuilder buf = new StringBuilder();
        XMLEvent xmlEvent;
        while((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            switch(xmlEvent.asStartElement().getName().getLocalPart()) {
                case "t": // Text
                    buf.append(xmlEventReader.getElementText());
                    break;
                case "r": // Rich Text Run
                    parseCT_RElt(xmlEventReader, buf);
                    break;
                case "rPh": // Phonetic Run
                case "phoneticPr": // Phonetic Properties
                    skipElement(xmlEventReader);
                    break;
                default:
                    throw new IllegalArgumentException("Unexpected start element: " + xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
        return buf.toString();
    }

    /**
     * Parses a {@code <r>} Rich Text Run. Returns just the text and drops the formatting. See <a
     * href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.run.aspx">xmlschema
     * type {@code CT_RElt}</a>.
     */
    static void parseCT_RElt(XMLEventReader xmlEventReader, StringBuilder buf) throws XMLStreamException {
        // Precondition: pointing to <r>;  Post condition: pointing to </r>
        XMLEvent xmlEvent;
        while((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            switch(xmlEvent.asStartElement().getName().getLocalPart()) {
                case "t": // Text
                    buf.append(xmlEventReader.getElementText());
                    break;
                case "rPr": // Run Properties
                    skipElement(xmlEventReader);
                    break;
                default:
                    throw new IllegalArgumentException("Unexpected start element: " + xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
    }

    static String getXMLText(XMLEventReader xmlEventReader, QName tag, List<String> wrappingTags) throws IOException, XMLStreamException {
        try (StringWriter sw = new StringWriter()) {
            XMLEventWriter xew = XMLHelper.newXMLOutputFactory().createXMLEventWriter(sw);
            try {
                for (String tagName : wrappingTags) {
                    xew.add(xef.createStartElement(new QName(NS_SPREADSHEETML, tagName),
                            Collections.emptyIterator(), Collections.emptyIterator()));
                }
                XMLEvent event = xmlEventReader.nextEvent();
                while (event != null && !(event.isEndElement() && event.asEndElement().getName().equals(tag))) {
                    xew.add(adjustNamespaceOnEvent(event));
                    event = xmlEventReader.nextEvent();
                }
                ListIterator<String> tagIter = wrappingTags.listIterator();
                while (tagIter.hasPrevious()) {
                    String tagName = tagIter.previous();
                    xew.add(xef.createEndElement(new QName(NS_SPREADSHEETML, tagName),
                            Collections.emptyIterator()));
                }
            } finally {
                xew.close();
            }
            return sw.toString();
        }
    }

    private static XMLEvent adjustNamespaceOnEvent(XMLEvent event) {
        if (event.isStartElement()) {
            StartElement se = event.asStartElement();
            String nsUri = se.getName().getNamespaceURI();
            if (nsUri != null && !nsUri.equals(NS_SPREADSHEETML)) {
                return xef.createStartElement(new QName(NS_SPREADSHEETML, se.getName().getLocalPart()),
                        se.getAttributes(), Collections.emptyIterator());
            }
        } else if (event.isEndElement()) {
            EndElement ee = event.asEndElement();
            String nsUri = ee.getName().getNamespaceURI();
            if (nsUri != null && !nsUri.equals(NS_SPREADSHEETML)) {
                return xef.createEndElement(new QName(NS_SPREADSHEETML, ee.getName().getLocalPart()),
                        Collections.emptyIterator());
            }
        }
        return event;
    }

    private static void skipElement(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to start element;  Post condition: pointing to end element
        while(xmlEventReader.nextTag().isStartElement()) {
            skipElement(xmlEventReader); // recursively skip over child
        }
    }
}
