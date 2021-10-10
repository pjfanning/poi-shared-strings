package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.util.XMLHelper;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLEventWriter;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;
import java.io.IOException;
import java.io.StringWriter;

class TextParser {

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

    static String getXMLText(XMLEventReader xmlEventReader, QName endTag) throws IOException, XMLStreamException {
        try (StringWriter sw = new StringWriter()) {
            XMLEventWriter xew = XMLHelper.newXMLOutputFactory().createXMLEventWriter(sw);
            try {
                XMLEvent event = xmlEventReader.nextEvent();
                while (event != null && !(event.isEndElement() && event.asEndElement().getName().equals(endTag))) {
                    xew.add(event);
                }
            } finally {
                xew.close();
            }
            return sw.toString();
        }
    }


    private static void skipElement(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to start element;  Post condition: pointing to end element
        while(xmlEventReader.nextTag().isStartElement()) {
            skipElement(xmlEventReader); // recursively skip over child
        }
    }
}
