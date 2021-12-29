package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;
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
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * Table of strings shared across all sheets in a workbook.
 * <p>
 * A workbook may contain thousands of cells containing string (non-numeric) data. Furthermore this data is very
 * likely to be repeated across many rows or columns. The goal of implementing a single string table that is shared
 * across the workbook is to improve performance in opening and saving the file by only reading and writing the
 * repetitive information once.
 * </p>
 * <p>
 * Consider for example a workbook summarizing information for cities within various countries. There may be a
 * column for the name of the country, a column for the name of each city in that country, and a column
 * containing the data for each city. In this case the country name is repetitive, being duplicated in many cells.
 * In many cases the repetition is extensive, and a tremendous savings is realized by making use of a shared string
 * table when saving the workbook. When displaying text in the spreadsheet, the cell table will just contain an
 * index into the string table as the value of a cell, instead of the full string.
 * </p>
 * <p>
 * The shared string table contains all the necessary information for displaying the string: the text, formatting
 * properties, and phonetic properties (for East Asian languages).
 * </p>
 */
public class TempFileSharedStringsTable extends SharedStringsTable {
    private static final Logger log = LoggerFactory.getLogger(TempFileSharedStringsTable.class);
    private static final QName COUNT_QNAME = new QName("count");
    private static final QName UNIQUE_COUNT_QNAME = new QName("uniqueCount");
    private File tempFile;
    private MVStore mvStore;
    private final boolean fullFormat;

    /**
     *  Array of individual string items in the Shared String table.
     */
    private final MVMap<Integer, String> strings;

    /**
     *  Maps strings and their indexes in the <code>strings</code> arrays
     */
    private final MVMap<String, Integer> stmap;

    private static final XmlOptions siSaveOptions = new XmlOptions(Constants.saveOptions);
    static {
        siSaveOptions.setSaveSyntheticDocumentElement(
                new QName(NS_SPREADSHEETML, "si"));
    }

    public TempFileSharedStringsTable() {
        this(false, false);
    }

    public TempFileSharedStringsTable(boolean encryptTempFiles) {
        this(encryptTempFiles, false);
    }

    public TempFileSharedStringsTable(boolean encryptTempFiles, boolean fullFormat) {
        super();
        this.fullFormat = fullFormat;
        try {
            tempFile = TempFile.createTempFile("poi-shared-strings", ".tmp");
            MVStore.Builder mvStoreBuilder = new MVStore.Builder();
            if (encryptTempFiles) {
                byte[] bytes = new byte[1024];
                Constants.RANDOM.nextBytes(bytes);
                mvStoreBuilder.encryptionKey(Base64.getEncoder().encodeToString(bytes).toCharArray());
            }
            mvStoreBuilder.fileName(tempFile.getAbsolutePath());
            mvStore = mvStoreBuilder.open();
            strings = mvStore.openMap("strings");
            stmap = mvStore.openMap("stmap");
        } catch (Error | RuntimeException e) {
            if (mvStore != null) mvStore.closeImmediately();
            if (tempFile != null) tempFile.delete();
            throw e;
        } catch (Exception e) {
            if (mvStore != null) mvStore.closeImmediately();
            if (tempFile != null) tempFile.delete();
            throw new RuntimeException(e);
        }
    }

    public TempFileSharedStringsTable(OPCPackage pkg, boolean encryptTempFiles) throws IOException {
        this(pkg, encryptTempFiles, false);
    }

    public TempFileSharedStringsTable(OPCPackage pkg, boolean encryptTempFiles,
                                      boolean fullFormat) throws IOException {
        this(encryptTempFiles, fullFormat);
        ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHARED_STRINGS.getContentType());
        if (parts.size() > 0) {
            PackagePart sstPart = parts.get(0);
            this.readFrom(sstPart.getInputStream());
        }
    }

    /**
     * Read this shared strings table from an XML file.
     * 
     * @param is The input stream containing the XML document.
     * @throws IOException if an error occurs while reading.
     */
    @Override
    public void readFrom(InputStream is) throws IOException {
        try {
            int uniqueCount = -1;
            int count = -1;
            XMLEventReader xmlEventReader = Constants.XML_INPUT_FACTORY.createXMLEventReader(is);
            try {
                while(xmlEventReader.hasNext()) {
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
                                addRSTEntry(new XSSFRichTextString(sst.getSiArray(0)).getCTRst(), true);
                            } else {
                                String text = TextParser.parseCT_Rst(xmlEventReader);
                                addPlainStringEntry(text, true);
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
        } catch(XMLStreamException e) {
            throw new IOException("Failed to parse shared strings", e);
        }
    }

    private CTRst getRSTEntryAt(int idx) throws XmlException, IOException {
        String str = strings.get(idx);
        if (str == null) throw new NoSuchElementException();
        return CTRst.Factory.parse(new StringReader(str));
    }

    private String getPlainStringEntryAt(int idx) {
        String str = strings.get(idx);
        if (str == null) throw new NoSuchElementException();
        return str;
    }

    /**
     * Return a string item by index
     *
     * @param idx index of item to return.
     * @return the item at the specified position in this Shared String table.
     * @throws NoSuchElementException if no item exists for this index
     * @throws POIXMLException if the item cannot be parsed
     */
    @Override
    public RichTextString getItemAt(int idx) {
        try {
            if (fullFormat) {
                return new XSSFRichTextString(getRSTEntryAt(idx));
            } else {
                return new XSSFRichTextString(getPlainStringEntryAt(idx));
            }
        } catch (NoSuchElementException nsee) {
            throw nsee;
        } catch (Exception e) {
            throw new POIXMLException("Failed to parse shared string", e);
        }
    }

    /**
     * Return an integer representing the total count of strings in the workbook. This count does not
     * include any numbers, it counts only the total of text strings in the workbook.
     *
     * @return the total count of strings in the workbook
     */
    @Override
    public int getCount() {
        return count;
    }

    /**
     * Returns an integer representing the total count of unique strings in the Shared String Table.
     * A string is unique even if it is a copy of another string, but has different formatting applied
     * at the character level.
     *
     * @return the total count of unique strings in the workbook
     */
    @Override
    public int getUniqueCount() {
        return uniqueCount;
    }

    private int addRSTEntry(CTRst st, boolean keepDuplicates) {
        if (st == null) {
            throw new NullPointerException("Cannot add null entry to SharedStringsTable");
        }
        String s = xmlText(st);
        count++;
        if (!keepDuplicates && stmap.containsKey(s)) {
            return stmap.get(s);
        }

        int idx = uniqueCount++;
        stmap.put(s, idx);
        strings.put(idx, st.xmlText());
        return idx;
    }

    private int addPlainStringEntry(String string, boolean keepDuplicates) {
        if (string == null) {
            throw new NullPointerException("Cannot add null entry to SharedStringsTable");
        }
        count++;
        if (!keepDuplicates && stmap.containsKey(string)) {
            return stmap.get(string);
        }

        int idx = uniqueCount++;
        stmap.put(string, idx);
        strings.put(idx, string);
        return idx;
    }

    /**
     * Add an entry to this Shared String table (a new value is appended to the end).
     *
     * <p>
     * If the Shared String table already contains this string entry, its index is returned.
     * Otherwise a new entry is added.
     * </p>
     *
     * @param string the entry to add
     * @return index the index of added entry
     */
    @Override
    public int addSharedStringItem(RichTextString string) {
        if(!(string instanceof XSSFRichTextString)){
            throw new IllegalArgumentException("Only XSSFRichTextString argument is supported");
        }
        if (fullFormat) {
            return addRSTEntry(((XSSFRichTextString) string).getCTRst(), false);
        } else {
            return addPlainStringEntry(string.getString(), false);
        }
    }

    /**
     * TempFileSharedStringsTable only supports streaming access of shared strings.
     * Use <code>getItemAt</code> instead
     *
     * @return list of shared string instances
     */
    @Override
    public List<RichTextString> getSharedStringItems() {
        throw new UnsupportedOperationException("TempFileSharedStringsTable only supports streaming access of shared strings");
    }

    /**
     * Write this table out as XML.
     * 
     * @param out The stream to write to.
     * @throws IOException if an error occurs while writing.
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
            Iterator<Integer> idIter = strings.keyIterator(null);
            while (idIter.hasNext()) {
                Integer stringId = idIter.next();
                XSSFRichTextString rst = (XSSFRichTextString)getItemAt(stringId);
                if (rst != null) {
                    writer.write(rst.getCTRst().xmlText(siSaveOptions));
                }
            }
            writer.write("</sst>");
        } finally {
            // do not close; let calling code close the output stream
            writer.flush();
        }
    }

    @Override
    public void close() throws IOException {
        if(mvStore != null) mvStore.closeImmediately();
        if(tempFile != null && !tempFile.delete()) {
            log.debug("failed to delete temp file - probably already deleted");
        }
    }
}
