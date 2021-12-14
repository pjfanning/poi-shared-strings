package com.github.pjfanning.poi.xssf.streaming.cache;

import com.github.pjfanning.poi.xssf.streaming.TextParser;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTSst;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.SstDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.Attribute;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import java.io.BufferedWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.OutputStreamWriter;
import java.io.Writer;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.NoSuchElementException;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

public class CachedSharedStringsTable extends SharedStringsTable {

    private static final Logger log = LoggerFactory.getLogger(CachedSharedStringsTable.class);
    private static final QName COUNT_QNAME = new QName("count");
    private static final QName UNIQUE_COUNT_QNAME = new QName("uniqueCount");
    private static final XmlOptions siSaveOptions = new XmlOptions(Constants.saveOptions);

    static {
        siSaveOptions.setSaveSyntheticDocumentElement(
                new QName(NS_SPREADSHEETML, "si"));
    }

    private final SSTCache sstCache;
    private final boolean fullFormat;

    public CachedSharedStringsTable(SSTCache sstCache, boolean fullFormat) {
        super();
        this.sstCache = sstCache;
        this.fullFormat = fullFormat;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void readFrom(InputStream is) throws IOException {
        try {
            int uniqueCount = -1;
            int count = -1;
            XMLEventReader xmlEventReader = Constants.XML_INPUT_FACTORY.createXMLEventReader(is);
            try {
                while (xmlEventReader.hasNext()) {
                    XMLEvent xmlEvent = xmlEventReader.nextEvent();

                    if (xmlEvent.isStartElement()) {
                        StartElement startElement = xmlEvent.asStartElement();
                        QName startTag = startElement.getName();
                        String localPart = startTag.getLocalPart();
                        if (localPart.equals("sst")) {
                            try {
                                Attribute countAtt = startElement.getAttributeByName(COUNT_QNAME);
                                if (countAtt != null) {
                                    count = Integer.parseInt(countAtt.getValue());
                                }
                            } catch (Exception e) {
                                log.warn("Failed to parse SharedStringsTable count");
                            }
                            try {
                                Attribute uniqueCountAtt = startElement.getAttributeByName(UNIQUE_COUNT_QNAME);
                                if (uniqueCountAtt != null) {
                                    uniqueCount = Integer.parseInt(uniqueCountAtt.getValue());
                                }
                            } catch (Exception e) {
                                log.warn("Failed to parse SharedStringsTable uniqueCount");
                            }
                        } else if (localPart.equals("si")) {
                            if (fullFormat) {
                                List<String> tags = Arrays.asList(new String[]{"sst", "si"});
                                String text = TextParser.getXMLText(xmlEventReader, startTag, tags);
                                CTSst sst;
                                try {
                                    sst = SstDocument.Factory.parse(text).getSst();
                                } catch (XmlException e) {
                                    throw new IOException("Failed to parse shared string text", e);
                                }
                                addEntry(new XSSFRichTextString(sst.getSiArray(0)).getCTRst(), true);
                            } else {
                                String text = TextParser.parseCT_Rst(xmlEventReader);
                                addEntry(new XSSFRichTextString(text).getCTRst(), true);
                            }
                        }
                    }
                }
                if (count > -1) {
                    this.count = count;
                }
                if (uniqueCount > -1) {
                    if (uniqueCount != this.uniqueCount) {
                        log.warn("SharedStringsTable has uniqueCount={} but read {} entries. This will probably cause some cells to be misinterpreted.",
                                uniqueCount, this.uniqueCount);
                    }
                    this.uniqueCount = uniqueCount;
                }
            } finally {
                xmlEventReader.close();
            }
        } catch (XMLStreamException e) {
            throw new IOException("Failed to parse shared strings", e);
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public RichTextString getItemAt(int idx) {
        return new XSSFRichTextString(getEntryAt(idx));
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getCount() {
        return count;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int getUniqueCount() {
        return uniqueCount;
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public int addSharedStringItem(RichTextString string) {
        if (!(string instanceof XSSFRichTextString)) {
            throw new IllegalArgumentException("Only XSSFRichTextString argument is supported");
        }
        return addEntry(((XSSFRichTextString) string).getCTRst(), false);
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public List<RichTextString> getSharedStringItems() {
        throw new UnsupportedOperationException("TempFileSharedStringsTable only supports streaming access of shared strings");
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void writeTo(OutputStream out) throws IOException {
        Writer writer = new BufferedWriter(new OutputStreamWriter(out, StandardCharsets.UTF_8));
        try {
            writer.write("<sst count=\"");
            writer.write(Integer.toString(count));
            writer.write("\" uniqueCount=\"");
            writer.write(Integer.toString(uniqueCount));
            writer.write("\" xmlns=\"");
            writer.write(NS_SPREADSHEETML);
            writer.write("\">");
            Iterator<Integer> idIterator = sstCache.keyIterator();
            while (idIterator.hasNext()) {
                CTRst rst = sstCache.getCTRst(idIterator.next());
                if (rst != null) {
                    writer.write(rst.xmlText(siSaveOptions));
                }
            }
            writer.write("</sst>");
        } finally {
            // do not close; let calling code close the output stream
            writer.flush();
        }
    }

    /**
     * {@inheritDoc}
     */
    @Override
    public void close() {
        sstCache.close();
    }

    private int addEntry(CTRst st, boolean keepDuplicates) {
        if (st == null) {
            throw new NullPointerException("Cannot add null entry to SharedStringsTable");
        }
        String s = xmlText(st);
        count++;
        if (!keepDuplicates && sstCache.containsString(s)) {
            return sstCache.getStringIndex(s);
        }

        int idx = uniqueCount++;
        sstCache.putStringIndex(s, idx);
        sstCache.putCTRst(idx, st);
        return idx;
    }

    private CTRst getEntryAt(int idx) {
        CTRst rst = sstCache.getCTRst(idx);
        if (rst == null) {
            throw new NoSuchElementException();
        }
        return rst;
    }

}
