package com.github.pjfanning.poi.xssf.streaming;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Base64;
import java.util.List;
import java.util.NoSuchElementException;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.util.StaxHelper;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.XMLEvent;

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
    private File tempFile;
    private MVStore mvStore;

    /**
     *  Array of individual string items in the Shared String table.
     */
    private final MVMap<Integer, CTRst> strings;

    /**
     *  Maps strings and their indexes in the <code>strings</code> arrays
     */
    private final MVMap<String, Integer> stmap;

    public TempFileSharedStringsTable() {
        this(false);
    }

    public TempFileSharedStringsTable(boolean encryptTempFiles) {
        super();
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
        this(encryptTempFiles);
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
            XMLEventReader xmlEventReader = StaxHelper.newXMLInputFactory().createXMLEventReader(is);

            while(xmlEventReader.hasNext()) {
                XMLEvent xmlEvent = xmlEventReader.nextEvent();

                if(xmlEvent.isStartElement() && xmlEvent.asStartElement().getName().getLocalPart().equals("si")) {
                    String str = parseCT_Rst(xmlEventReader);
                    addSharedStringItem(new XSSFRichTextString(str));
                }
            }
        } catch(XMLStreamException e) {
            throw new IOException(e);
        }
    }

    /**
     * Return a string item by index
     *
     * @param idx index of item to return.
     * @return the item at the specified position in this Shared String table.
     * @deprecated use <code>getItemAt(int idx)</code> instead
     * @throws NoSuchElementException if no item exists for this index
     */
    @Deprecated
    @Override
    public CTRst getEntryAt(int idx) {
        CTRst rst = strings.get(idx);
        if (rst == null) throw new NoSuchElementException();
        return rst;
    }

    /**
     * Return a string item by index
     *
     * @param idx index of item to return.
     * @return the item at the specified position in this Shared String table.
     * @throws NoSuchElementException if no item exists for this index
     */
    @Override
    public RichTextString getItemAt(int idx) {
        return new XSSFRichTextString(getEntryAt(idx));
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

    /**
     * Add an entry to this Shared String table (a new value is appended to the end).
     *
     * <p>
     * If the Shared String table already contains this <code>CTRst</code> bean, its index is returned.
     * Otherwise a new entry is aded.
     * </p>
     *
     * @param st the entry to add
     * @return index the index of added entry
     * @deprecated use <code>addSharedStringItem(RichTextString string)</code> instead
     */
    @Deprecated
    @Override
    public int addEntry(CTRst st) {
        if (st == null) {
            throw new NullPointerException("Cannot add null entry to SharedStringsTable");
        }
        String s = xmlText(st);
        count++;
        if (stmap.containsKey(s)) {
            return stmap.get(s);
        }

        int idx = uniqueCount++;
        stmap.put(s, idx);
        strings.put(idx, st);
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
     * @since POI 4.0.0
     * @return index the index of added entry
     */
    @Override
    public int addSharedStringItem(RichTextString string) {
        if(!(string instanceof XSSFRichTextString)){
            throw new IllegalArgumentException("Only XSSFRichTextString argument is supported");
        }
        return addEntry(((XSSFRichTextString) string).getCTRst());
    }

    /**
     * TempFileSharedStringsTable only supports streaming access of shared strings
     *
     * @return array of CTRst beans
     * @deprecated use <code>getItemAt</code> instead
     */
    @Deprecated
    @Override
    public List<CTRst> getItems() {
        throw new UnsupportedOperationException("TempFileSharedStringsTable only supports streaming access of shared strings");
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
            for (CTRst rst : strings.values()) {
                writer.write("<si>");
                writer.write(xmlText(rst));
                writer.write("</si>");
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
        if(tempFile != null) tempFile.delete();
    }

    /**
     * Parses a {@code <si>} String Item. Returns just the text and drops the formatting. See <a
     * href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sharedstringitem.aspx">xmlschema
     * type {@code CT_Rst}</a>.
     */
    private String parseCT_Rst(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to <si>;  Post condition: pointing to </si>
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
                    throw new IllegalArgumentException(xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
        return buf.toString();
    }

    /**
     * Parses a {@code <r>} Rich Text Run. Returns just the text and drops the formatting. See <a
     * href="https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.run.aspx">xmlschema
     * type {@code CT_RElt}</a>.
     */
    private void parseCT_RElt(XMLEventReader xmlEventReader, StringBuilder buf) throws XMLStreamException {
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
                    throw new IllegalArgumentException(xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
    }

    private void skipElement(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to start element;  Post condition: pointing to end element
        while(xmlEventReader.nextTag().isStartElement()) {
            skipElement(xmlEventReader); // recursively skip over child
        }
    }
}
