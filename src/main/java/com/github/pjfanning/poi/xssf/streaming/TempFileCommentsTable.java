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
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCommentList;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CommentsDocument;
import org.xml.sax.SAXException;

import javax.xml.namespace.QName;
import javax.xml.parsers.ParserConfigurationException;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.events.StartElement;
import javax.xml.stream.events.XMLEvent;
import javax.xml.transform.TransformerException;
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

    private final boolean fullFormat;
    private final MVMap<String, XSSFComment> comments;
    private final MVMap<Integer, String> authors;

    private static final XmlOptions options = new XmlOptions();
    static {
        options.setSaveInner();
        options.setSaveAggressiveNamespaces();
        options.setUseDefaultNamespace(true);
        options.setSaveUseOpenFrag(false);
        options.setSaveImplicitNamespaces(Collections.singletonMap("", NS_SPREADSHEETML));
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
            XMLEventReader xmlEventReader = XMLHelper.newXMLInputFactory().createXMLEventReader(is);
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
                            XSSFComment xc = new SimpleXSSFComment();
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
                writer.write("\">");
                XSSFRichTextString rts = comment.getString();
                if (rts != null) {
                    writer.write("<text>");
                    if (rts.getCTRst() != null) {
                        writer.write(WriteUtils.stripXmlFragmentElement(rts.getCTRst().xmlText(options)));
                    } else {
                        writer.write("<t>");
                        writer.write(StringEscapeUtils.escapeXml11(comment.getString().getString()));
                        writer.write("</t>");
                    }
                    writer.write("</text>");
                }
                writer.write("</comment>");
            }
            writer.write("</commentList>");
            writer.write("</comments>");
        } catch (SAXException | ParserConfigurationException | TransformerException e) {
            throw new IOException("Problem writing comments data", e);
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
                    throw new IllegalArgumentException("Unexpected element name " + xmlEvent.asStartElement().getName().getLocalPart());
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
            switch(xmlEvent.asStartElement().getName().getLocalPart()) {
                case "text":
                    text = TextParser.parseCT_Rst(xmlEventReader);
                    break;
                default:
                    throw new IllegalArgumentException("Unexpected element name " + xmlEvent.asStartElement().getName().getLocalPart());
            }
        }
        return text;
    }
}
