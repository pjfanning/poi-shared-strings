package com.github.pjfanning.poi.xssf.streaming;

import com.microsoft.schemas.vml.CTShape;
import org.apache.commons.text.StringEscapeUtils;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.Internal;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFVMLDrawing;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTCommentList;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CommentsDocument;
import org.slf4j.Logger;

import javax.xml.namespace.QName;
import javax.xml.stream.XMLEventReader;
import javax.xml.stream.XMLStreamException;
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
import java.util.concurrent.ConcurrentMap;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

/**
 * Table of comments.
 * <p>
 * The comments table contains all the necessary information for displaying the string: the text, formatting
 * properties, and phonetic properties (for East Asian languages).
 * </p>
 */
public abstract class CommentsTableBase extends POIXMLDocumentPart implements Comments, AutoCloseable {
    protected Sheet sheet;
    protected boolean ignoreDrawing = false;
    protected final boolean fullFormat;
    protected ConcurrentMap<String, SerializableComment> comments;
    protected ConcurrentMap<Integer, String> authors;

    private static final XmlOptions textSaveOptions = new XmlOptions(Constants.saveOptions);
    static {
        textSaveOptions.setSaveSyntheticDocumentElement(
                new QName(NS_SPREADSHEETML, "text"));
    }

    protected CommentsTableBase(boolean fullFormat) {
        super();
        this.fullFormat = fullFormat;
    }

    protected abstract Logger getLogger();

    protected abstract Iterator<Integer> authorsKeyIterator();

    protected abstract Iterator<String> commentsKeyIterator();

    /**
     * @param ignoreDrawing set to true if you don't need the drawing/shape data on the comments
     *                      (default is false) - ignoring the drawing/shape data can save memory
     */
    public void setIgnoreDrawing(boolean ignoreDrawing) {
        this.ignoreDrawing = ignoreDrawing;
    }

    /**
     * @return whether to ignore the drawing/shape data (default is false) -
     *         ignoring the drawing/shape data can save memory
     */
    public boolean isIgnoreDrawing() {
        return ignoreDrawing;
    }

    @Override
    @Internal
    public void setSheet(Sheet sheet) {
        this.sheet = sheet;
    }

    @Override
    protected void commit() throws IOException {
        PackagePart part = getPackagePart();
        try (OutputStream out = part.getOutputStream()) {
            writeTo(out);
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
        String nullSafeAuthor = author == null ? "" : author;
        Iterator<Integer> authorIdIterator = authorsKeyIterator();
        while (authorIdIterator.hasNext()) {
            Integer authorId = authorIdIterator.next();
            String existingAuthor = authorId == null ? null : authors.get(authorId);
            if (nullSafeAuthor.equals(existingAuthor)) {
                return authorId;
            }
        }
        int index = getNumberOfAuthors();
        if (index == 0 && !nullSafeAuthor.equals("")) {
            authors.put(index++, "");
        }
        authors.put(index, nullSafeAuthor);
        return index;
    }

    @Override
    public XSSFComment findCellComment(CellAddress cellAddress) {
        SerializableComment serializableComment = comments.get(cellAddress.formatAsString());
        if (serializableComment == null) {
            return null;
        }
        XSSFVMLDrawing vml = getVMLDrawing(sheet, false);
        return new DelegatingXSSFComment(this, serializableComment,
                vml == null ? null : vml.findCommentShape(cellAddress.getRow(), cellAddress.getColumn()));
    }

    @Override
    public boolean removeComment(CellAddress cellRef) {
        return comments.remove(cellRef.formatAsString()) != null;
    }

    @Override
    public Iterator<CellAddress> getCellAddresses() {
        final Iterator<String> keyIterator = commentsKeyIterator();
        return new Iterator<CellAddress>() {
            @Override
            public boolean hasNext() {
                return keyIterator.hasNext();
            }

            @Override
            public CellAddress next() {
                return new CellAddress(keyIterator.next());
            }
        };
    }

    @Override
    public XSSFComment createNewComment(ClientAnchor clientAnchor) {
        XSSFVMLDrawing vml = getVMLDrawing(sheet, true);
        CTShape vmlShape = vml == null ? null : vml.newCommentShape();
        if (vmlShape != null && clientAnchor instanceof XSSFClientAnchor && ((XSSFClientAnchor)clientAnchor).isSet()) {
            // convert offsets from emus to pixels since we get a
            // DrawingML-anchor
            // but create a VML Drawing
            int dx1Pixels = clientAnchor.getDx1() / Units.EMU_PER_PIXEL;
            int dy1Pixels = clientAnchor.getDy1() / Units.EMU_PER_PIXEL;
            int dx2Pixels = clientAnchor.getDx2() / Units.EMU_PER_PIXEL;
            int dy2Pixels = clientAnchor.getDy2() / Units.EMU_PER_PIXEL;
            String position = clientAnchor.getCol1() + ", " + dx1Pixels + ", " + clientAnchor.getRow1() + ", " + dy1Pixels + ", " +
                    clientAnchor.getCol2() + ", " + dx2Pixels + ", " + clientAnchor.getRow2() + ", " + dy2Pixels;
            vmlShape.getClientDataArray(0).setAnchorArray(0, position);
        }
        CellAddress ref = new CellAddress(clientAnchor.getRow1(), clientAnchor.getCol1());

        if (findCellComment(ref) != null) {
            throw new IllegalArgumentException("Multiple cell comments in one cell are not allowed, cell: " + ref);
        }

        String key = ref.formatAsString();
        CTComment ctComment = CTComment.Factory.newInstance();
        ctComment.setRef(key);
        SerializableComment serializableComment = new SerializableComment();
        serializableComment.setAddress(ref);
        comments.put(key, serializableComment);

        return new XSSFComment(this, ctComment, vmlShape);
    }

    @Override
    public void referenceUpdated(CellAddress oldReference, XSSFComment comment) {
        removeComment(oldReference);
        addToMap(comment);
    }

    @Override
    public void commentUpdated(XSSFComment comment) {
        removeComment(comment.getAddress());
        addToMap(comment);
    }

    private void addToMap(XSSFComment comment) {
        SerializableComment serializableComment = new SerializableComment();
        serializableComment.setAddress(comment.getAddress());
        serializableComment.setString(comment.getString());
        serializableComment.setAuthor(comment.getAuthor());
        serializableComment.setVisible(comment.isVisible());
        comments.put(comment.getAddress().formatAsString(), serializableComment);
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
            Iterator<Integer> authorIdIterator = authorsKeyIterator();
            while (authorIdIterator.hasNext()) {
                Integer authorId = authorIdIterator.next();
                String author = authorId == null ? null : authors.get(authorId);
                writer.write("<author>");
                writer.write(StringEscapeUtils.escapeXml11(author));
                writer.write("</author>");
            }
            writer.write("</authors>");
            writer.write("<commentList>");
            Iterator<String> commentsRefIterator = commentsKeyIterator();
            while (commentsRefIterator.hasNext()) {
                SerializableComment comment = comments.get(commentsRefIterator.next());
                if (comment != null) {
                    writer.write("<comment ref=\"");
                    writer.write(StringEscapeUtils.escapeXml11(comment.getAddress().formatAsString()));
                    String author = comment.getAuthor();
                    int authorId = findAuthor(author);
                    writer.write("\" authorId=\"");
                    writer.write(Integer.toString(authorId));
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
        while((xmlEvent = xmlEventReader.nextEvent()).isStartElement()) {
            StartElement startElement = xmlEvent.asStartElement();
            QName startTag = startElement.getName();
            switch(startTag.getLocalPart()) {
                case "text":
                    List<String> tags = Arrays.asList(new String[]{"comments", "commentList", "comment", "text"});
                    String text = TextParser.getXMLText(xmlEventReader, startTag, tags);
                    CTCommentList commentsList = CommentsDocument.Factory.parse(text).getComments().getCommentList();
                    richTextString = new XSSFRichTextString(commentsList.getCommentArray(0).getText());
                    break;
                default:
                    getLogger().debug("ignoring data inside element {}", startElement.getName());
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
        while((xmlEvent = xmlEventReader.nextEvent()).isStartElement()) {
            StartElement startElement = xmlEvent.asStartElement();
            switch(startElement.getName().getLocalPart()) {
                case "text":
                    text = TextParser.parseCT_Rst(xmlEventReader);
                    break;
                default:
                    getLogger().debug("ignoring data inside element {}", startElement.getName());
                    break;
            }
        }
        return text;
    }

    private XSSFVMLDrawing getVMLDrawing(Sheet sheet, boolean autocreate) {
        if (!ignoreDrawing) {
            if (sheet instanceof XSSFSheet) {
                return ((XSSFSheet)sheet).getVMLDrawing(autocreate);
            } else if (sheet instanceof SXSSFSheet) {
                return ((SXSSFSheet)sheet).getVMLDrawing(autocreate);
            }
        }
        return null;
    }
}
