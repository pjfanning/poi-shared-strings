package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.StaxHelper;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Base64;
import java.util.Iterator;
import java.util.Set;
import java.util.stream.Collectors;

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
public class TempFileCommentsTable implements Comments {
    private File tempFile;
    private MVStore mvStore;

    private final MVMap<String, XSSFComment> comments;

    private final MVMap<Integer, String> authors;

    public TempFileCommentsTable() {
        this(false);
    }

    public TempFileCommentsTable(boolean encryptTempFiles) {
        super();
        try {
            tempFile = TempFile.createTempFile("poi-comments", ".tmp");
            MVStore.Builder mvStoreBuilder = new MVStore.Builder();
            if (encryptTempFiles) {
                byte[] bytes = new byte[1024];
                Constants.RANDOM.nextBytes(bytes);
                mvStoreBuilder.encryptionKey(Base64.getEncoder().encodeToString(bytes).toCharArray());
            }
            mvStoreBuilder.fileName(tempFile.getAbsolutePath());
            mvStore = mvStoreBuilder.open();
            comments = mvStore.openMap("comments");
            authors = mvStore.openMap("authors");
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

    public TempFileCommentsTable(OPCPackage pkg, boolean encryptTempFiles) throws IOException {
        this(encryptTempFiles);
        ArrayList<PackagePart> parts = pkg.getPartsByContentType(XSSFRelation.SHEET_COMMENTS.getContentType());
        if (parts.size() > 0) {
            PackagePart sstPart = parts.get(0);
            this.readFrom(sstPart.getInputStream());
        }
    }

    /**
     * Read this comments table from an XML file.
     * 
     * @param is The input stream containing the XML document.
     * @throws IOException if an error occurs while reading.
     */
    public void readFrom(InputStream is) throws IOException {
        try {
            XMLEventReader xmlEventReader = StaxHelper.newXMLInputFactory().createXMLEventReader(is);

            while(xmlEventReader.hasNext()) {
                XMLEvent xmlEvent = xmlEventReader.nextEvent();

                if(xmlEvent.isStartElement()) {
                    StartElement se = xmlEvent.asStartElement();
                    if(se.getName().getLocalPart().equals("author")) {
                        authors.put(getNumberOfAuthors(), xmlEventReader.getElementText());
                    } else if(se.getName().getLocalPart().equals("comment")) {
                        String ref = se.getAttributeByName(new QName("ref")).getValue();
                        String authorId = se.getAttributeByName(new QName("authorId")).getValue();
                        String str = parseComment(xmlEventReader);
                        XSSFComment xc = new XSSFComment(null, null, null);
                        xc.setAddress(new CellAddress(ref));
                        xc.setAuthor(authors.get(authorId));
                        xc.setString(str);
                        comments.put(ref, xc);
                    }
                }
            }
        } catch(XMLStreamException e) {
            throw new IOException(e);
        }
    }

    @Override
    public int getNumberOfComments() {
        return comments.size();
    }

    @Override
    public int getNumberOfAuthors() {
        return authors.size();
    }

    @Override
    public String getAuthor(long authorId) {
        return authors.get(Long.toString(authorId));
    }

    @Override
    public int findAuthor(String author) {
        return 0;
    }

    @Override
    public XSSFComment findCellComment(CellAddress cellAddress) {
        return comments.get(cellAddress.formatAsString());
    }

    @Override
    public boolean removeComment(CellAddress cellRef) {
        return false;
    }

    @Override
    public Iterator<CellAddress> getCellAddresses() {
        Set<String> set = comments.keySet();
        return set.stream().map((s) -> new CellAddress(s)).collect(Collectors.toSet()).iterator();
    }

    /**
     * Parses a {@code <comment>} Comment. Returns just the text and drops the formatting.
     */
    private String parseComment(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to <comment>;  Post condition: pointing to </comment>
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
