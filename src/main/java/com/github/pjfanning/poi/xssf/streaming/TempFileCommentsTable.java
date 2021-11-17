package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCommentList;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CommentsDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

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
 * This implementation does not extend <code>CommentsTable</code>, so cannot be used for
 * creating new comments - it can only be used for reading existing comments from saved files.
 */
public class TempFileCommentsTable extends POIXMLDocumentPart implements Comments, AutoCloseable {
    private static Logger log = LoggerFactory.getLogger(TempFileCommentsTable.class);

    private File tempFile;
    private MVStore mvStore;

    private final boolean fullFormat;
    private final MVMap<String, SerializableComment> comments;
    private final MVMap<Integer, String> authors;

    private static final XmlOptions textSaveOptions = new XmlOptions(Constants.saveOptions);
    static {
        textSaveOptions.setSaveSyntheticDocumentElement(
                new QName(NS_SPREADSHEETML, "text"));
    }

    public TempFileCommentsTable() {
        this(false, false);
    }

    public TempFileCommentsTable(boolean encryptTempFiles) {
        this(encryptTempFiles, false);
    }

    public TempFileCommentsTable(boolean encryptTempFiles, boolean fullFormat) {
        super();
        this.fullFormat = fullFormat;
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
        this(pkg, encryptTempFiles, false);
    }

    public TempFileCommentsTable(OPCPackage pkg, boolean encryptTempFiles,
                                 boolean fullFormat) throws IOException {
        this(encryptTempFiles, fullFormat);
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
            XMLEventReader xmlEventReader = Constants.XML_INPUT_FACTORY.createXMLEventReader(is);
            try {
                while(xmlEventReader.hasNext()) {
                    XMLEvent xmlEvent = xmlEventReader.nextEvent();

                    if (xmlEvent.isStartElement()) {
                        StartElement se = xmlEvent.asStartElement();
                        if (se.getName().getLocalPart().equals("author")) {
                            authors.put(getNumberOfAuthors(), xmlEventReader.getElementText());
                        } else if (se.getName().getLocalPart().equals("comment")) {
                            String ref = se.getAttributeByName(new QName("ref")).getValue();
                            String authorId = se.getAttributeByName(new QName("authorId")).getValue();
                            XSSFRichTextString str;
                            if (fullFormat) {
                                try {
                                    str = parseFullComment(xmlEventReader);
                                } catch (XmlException e) {
                                    throw new IOException("Failed to parse comment", e);
                                }
                            } else {
                                str = new XSSFRichTextString(parseSimplifiedComment(xmlEventReader));
                            }
                            SerializableComment xc = new SerializableComment();
                            xc.setAddress(new CellAddress(ref));
                            xc.setAuthor(authors.get(Integer.parseInt(authorId)));
                            xc.setString(str);
                            comments.put(ref, xc);
                        }
                    }
                }
            } finally {
                xmlEventReader.close();
            }
        } catch (XMLStreamException xse) {
            throw new IOException("Failed to parse comments", xse);
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
        SerializableComment comment = comments.get(cellAddress.formatAsString());
        return comment == null ? null : new ReadOnlyXSSFComment(comment);
    }

    /**
     * Not implemented. This class only supports read-only view of Comments.
     * @throws IllegalStateException
     */
    @Override
    public boolean removeComment(CellAddress cellRef) {
        throw new IllegalStateException("Not Implemented - this class only supports read-only view of Comments");
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
            for (SerializableComment comment : comments.values()) {
                writer.write("<comment ref=\"");
                writer.write(StringEscapeUtils.escapeXml11(comment.getAddress().formatAsString()));
                writer.write("\" authorId=\"");
                writer.write(Integer.toString(findAuthor(comment.getAuthor())));
                writer.write("\">");
                XSSFRichTextString rts = comment.getString();
                if (rts != null) {
                    if (rts.getCTRst() != null) {
                        writer.write(rts.getCTRst().xmlText(textSaveOptions));
                    } else {
                        writer.write("<text><t>");
                        writer.write(StringEscapeUtils.escapeXml11(comment.getString().getString()));
                        writer.write("</t></text>");
                    }
                }
                writer.write("</comment>");
            }
            writer.write("</commentList>");
            writer.write("</comments>");
        } finally {
            // do not close; let calling code close the output stream
            writer.flush();
        }
    }

    /**
     * Parses a {@code <comment>} Comment. Uses POI/XMLBeans classes to parse full comment XML.
     */
    private XSSFRichTextString parseFullComment(XMLEventReader xmlEventReader) throws IOException, XmlException, XMLStreamException {
        // Precondition: pointing to <comment>;  Post condition: pointing to </comment>
        XMLEvent xmlEvent;
        XSSFRichTextString richTextString = null;
        while((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            StartElement startElement = xmlEvent.asStartElement();
            QName startTag = startElement.getName();
            switch(startTag.getLocalPart()) {
                case "text":
                    List<String> tags = Arrays.asList(new String[]{"comments", "commentList", "comment", "text"});
                    String text = TextParser.getXMLText(xmlEventReader, startTag, tags);
                    CTCommentList comments = CommentsDocument.Factory.parse(text).getComments().getCommentList();
                    richTextString = new XSSFRichTextString(comments.getCommentArray(0).getText());
                    break;
                default:
                    log.debug("ignoring data inside element {}", startElement.getName());
                    break;
            }
        }
        return richTextString;
    }

    /**
     * Parses a {@code <comment>} Comment. Returns just the text and drops the formatting.
     */
    private String parseSimplifiedComment(XMLEventReader xmlEventReader) throws XMLStreamException {
        // Precondition: pointing to <comment>;  Post condition: pointing to </comment>
        XMLEvent xmlEvent;
        String text = null;
        while((xmlEvent = xmlEventReader.nextTag()).isStartElement()) {
            StartElement startElement = xmlEvent.asStartElement();
            switch(startElement.getName().getLocalPart()) {
                case "text":
                    text = TextParser.parseCT_Rst(xmlEventReader);
                    break;
                default:
                    log.debug("ignoring data inside element {}", startElement.getName());
                    break;
            }
        }
        return text;
    }
}
