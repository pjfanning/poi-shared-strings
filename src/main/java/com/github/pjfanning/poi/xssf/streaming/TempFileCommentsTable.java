package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.TempFile;
import org.apache.poi.util.XMLHelper;
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
import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.*;
import java.util.stream.Collectors;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * Table of comments.
 * <p>
 * The comments table contains all the necessary information for displaying the string: the text, formatting
 * properties, and phonetic properties (for East Asian languages).
 * </p>
 */
public class TempFileCommentsTable extends POIXMLDocumentPart implements Comments, AutoCloseable {
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
            XMLEventReader xmlEventReader = XMLHelper.newXMLInputFactory().createXMLEventReader(is);

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
                        XSSFComment xc = new SimpleXSSFComment();
                        xc.setAddress(new CellAddress(ref));
                        xc.setAuthor(authors.get(Integer.parseInt(authorId)));
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
        return authors.get((int)authorId);
    }

    @Override
    public int findAuthor(String author) {
        for (Map.Entry<Integer, String> entry : authors.entrySet()) {
            if (entry.getValue().equals(author)) {
                return entry.getKey();
            }
        }
        int index = getNumberOfAuthors();
        authors.put(index, author);
        return index;
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

    @Override
    public void close() {
        if(mvStore != null) mvStore.closeImmediately();
        if(tempFile != null) tempFile.delete();
    }

    /**
     * Write this table out as XML.
     *
     * @param out The stream to write to.
     * @throws IOException if an error occurs while writing.
     */
    public void writeTo(OutputStream out) throws IOException {
        Writer writer = new BufferedWriter(new OutputStreamWriter(out, StandardCharsets.UTF_8));
        try {
            writer.write("<comments xmlns=\"");
            writer.write(NS_SPREADSHEETML);
            writer.write("\"><authors>");
            for (String author : authors.values()) {
                writer.write("<author>");
                writer.write(StringEscapeUtils.escapeXml11(author));
                writer.write("</author>");
            }
            writer.write("</authors>");
            writer.write("<commentList>");
            for (XSSFComment comment : comments.values()) {
                writer.write("<comment ref=\"");
                writer.write(StringEscapeUtils.escapeXml11(comment.getAddress().formatAsString()));
                writer.write("\" authorId=\"");
                writer.write(Integer.toString(findAuthor(comment.getAuthor())));
                writer.write("\"><text><t>");
                writer.write(StringEscapeUtils.escapeXml11(comment.getString().getString()));
                writer.write("</t></text></comment>");
            }
            writer.write("</commentList>");
            writer.write("</comments>");
        } finally {
            // do not close; let calling code close the output stream
            writer.flush();
        }
    }

    /**
     * Parses a {@code <comment>} Comment. Returns just the text and drops the formatting.
     */
    private String parseComment(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to <comment>;  Post condition: pointing to </comment>
        XMLEvent xmlEvent;
        String text = null;
        while((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            switch(xmlEvent.asStartElement().getName().getLocalPart()) {
                case "text":
                    text = TextParser.parseCT_Rst(xmlEventReader);
                    break;
                default:
                    throw new IllegalArgumentException(xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
        return text;
    }
}
