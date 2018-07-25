package com.github.pjfanning.poi.xssf.streaming;

import static org.apache.poi.xssf.usermodel.XSSFRelation.NS_SPREADSHEETML;

import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.util.*;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.h2.mvstore.MVMap;
import org.h2.mvstore.MVStore;
import org.xml.sax.Attributes;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

/**
 * <p>This is a lightweight way to process the Shared Strings
 *  table. Most of the text cells will reference something
 *  from in here.
 * <p>Note that each SI entry can have multiple T elements, if the
 *  string is made up of bits with different formatting.
 * <p>Example input:
 * <pre>
 &lt;?xml version="1.0" encoding="UTF-8" standalone="yes" ?>
 &lt;sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">
 &lt;si>
 &lt;r>
 &lt;rPr>
 &lt;b />
 &lt;sz val="11" />
 &lt;color theme="1" />
 &lt;rFont val="Calibri" />
 &lt;family val="2" />
 &lt;scheme val="minor" />
 &lt;/rPr>
 &lt;t>This:&lt;/t>
 &lt;/r>
 &lt;r>
 &lt;rPr>
 &lt;sz val="11" />
 &lt;color theme="1" />
 &lt;rFont val="Calibri" />
 &lt;family val="2" />
 &lt;scheme val="minor" />
 &lt;/rPr>
 &lt;t xml:space="preserve">Causes Problems&lt;/t>
 &lt;/r>
 &lt;/si>
 &lt;si>
 &lt;t>This does not&lt;/t>
 &lt;/si>
 &lt;/sst>
 * </pre>
 *
 */
public class ReadOnlyTempFileSharedStringsTable extends ReadOnlySharedStringsTable implements Closeable {

    private File tempFile;
    private MVStore mvStore;
    private final boolean encryptTempFiles = true;

    /**
     * The shared strings table.
     */
    private MVMap<Integer, String> strings;

    private int index;

    /**
     * Calls {{@link #ReadOnlyTempFileSharedStringsTable(OPCPackage, boolean)}} with
     * a value of <code>true</code> for including phonetic runs
     *
     * @param pkg The {@link OPCPackage} to use as basis for the shared-strings table.
     * @throws IOException If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    public ReadOnlyTempFileSharedStringsTable(OPCPackage pkg)
            throws IOException, SAXException {
        this(pkg, true);
    }

    /**
     *
     * @param pkg The {@link OPCPackage} to use as basis for the shared-strings table.
     * @param includePhoneticRuns whether or not to concatenate phoneticRuns onto the shared string
     * @since POI 3.14-Beta3
     * @throws IOException If reading the data from the package fails.
     * @throws SAXException if parsing the XML data fails.
     */
    public ReadOnlyTempFileSharedStringsTable(OPCPackage pkg, boolean includePhoneticRuns)
            throws IOException, SAXException {
        super(pkg, includePhoneticRuns);
    }

    /**
     * Read this shared strings table from an XML file.
     *
     * @param is The input stream containing the XML document.
     * @throws IOException if an error occurs while reading.
     * @throws SAXException if parsing the XML data fails.
     */
    @Override
    public void readFrom(InputStream is) throws IOException, SAXException {
        // test if the file is empty, otherwise parse it
        PushbackInputStream pis = new PushbackInputStream(is, 1);
        int emptyTest = pis.read();
        if (emptyTest > -1) {
            pis.unread(emptyTest);
            InputSource sheetSource = new InputSource(pis);
            try {
                XMLReader sheetParser = SAXHelper.newXMLReader();
                sheetParser.setContentHandler(this);
                sheetParser.parse(sheetSource);
            } catch(ParserConfigurationException e) {
                throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
            }
        }
    }

    /**
     * Return the string at a given index.
     * Formatting is ignored.
     *
     * @param idx index of item to return.
     * @return the item at the specified position in this Shared String table.
     */
    @Override
    public String getEntryAt(int idx) {
        return strings.get(idx);
    }

    @Override
    public List<String> getItems() {
        return new ArrayList<>(strings.values());
    }

    @Override
    public void close() throws IOException {
        if(mvStore != null) mvStore.closeImmediately();
        if(tempFile != null) tempFile.delete();
    }

    //// ContentHandler methods ////

    private StringBuilder characters;
    private boolean tIsOpen;
    private boolean inRPh;

    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {
        if (uri != null && ! uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        if ("sst".equals(localName)) {
            String count = attributes.getValue("count");
            if(count != null) this.count = Integer.parseInt(count);
            String uniqueCount = attributes.getValue("uniqueCount");
            if(uniqueCount != null) this.uniqueCount = Integer.parseInt(uniqueCount);

            characters = new StringBuilder(64);
        } else if ("si".equals(localName)) {
            characters.setLength(0);
        } else if ("t".equals(localName)) {
            tIsOpen = true;
        } else if ("rPh".equals(localName)) {
            inRPh = true;
            //append space...this assumes that rPh always comes after regular <t>
            if (includePhoneticRuns && characters.length() > 0) {
                characters.append(" ");
            }
        }
    }

    public void endElement(String uri, String localName, String name)
            throws SAXException {
        if (uri != null && ! uri.equals(NS_SPREADSHEETML)) {
            return;
        }

        if ("si".equals(localName)) {
            if (strings == null) {
                initMap();
            }
            strings.put(index++, characters.toString());
        } else if ("t".equals(localName)) {
            tIsOpen = false;
        } else if ("rPh".equals(localName)) {
            inRPh = false;
        }
    }

    /**
     * Captures characters only if a t(ext) element is open.
     */
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        if (tIsOpen) {
            if (inRPh && includePhoneticRuns) {
                characters.append(ch, start, length);
            } else if (! inRPh){
                characters.append(ch, start, length);
            }
        }
    }

    private void initMap() {
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
}
