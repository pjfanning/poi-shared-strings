package com.github.pjfanning.poi.xssf.streaming;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.UUID;
import java.util.regex.Pattern;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Test;
import org.xml.sax.SAXException;

import static com.github.pjfanning.poi.xssf.streaming.TestIOUtils.getResourceStream;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertThrows;

public class TestTempFileSharedStringsTable {
    @Test
    public void testWriteOut() throws Exception {
        testWriteOut(false);
    }

    @Test
    public void testWriteOutFullFormat() throws Exception {
        testWriteOut(true);
    }

    @Test
    public void testReadXML() throws Exception {
        testReadXML(false, false);
    }

    @Test
    public void testReadXMLEncrypted() throws Exception {
        testReadXML(true, false);
    }

    @Test
    public void testReadXMLFullFormat() throws Exception {
        testReadXML(false, true);
    }

    @Test
    public void testReadXMLEncryptedFullFormat() throws Exception {
        testReadXML(true, true);
    }

    @Test
    public void testReadXMLWithPhoneticHints() throws Exception {
        try (InputStream is = getResourceStream("sharedStrings-with-phonetic-hints.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(false, false)) {
            sst.readFrom(is);
            assertEquals(3, sst.getUniqueCount());
            assertEquals(3, sst.getCount());
            assertEquals("Country", sst.getItemAt(0).getString());
            assertEquals("Country", sst.getString(0));
            assertEquals("City", sst.getItemAt(1).getString());
            assertEquals("City", sst.getString(1));
            assertEquals("沖縄", sst.getItemAt(2).getString());
            assertEquals("沖縄", sst.getString(2));
        }
    }

    @Test
    public void testReadXMLWithPhoneticHintsPOISST() throws Exception {
        try (InputStream is = getResourceStream("sharedStrings-with-phonetic-hints.xml");
             SharedStringsTable sst = new SharedStringsTable()) {
            sst.readFrom(is);
            assertEquals(3, sst.getUniqueCount());
            assertEquals(3, sst.getCount());
            assertEquals("Country", sst.getItemAt(0).getString());
            assertEquals("City", sst.getItemAt(1).getString());
            assertEquals("沖縄", sst.getItemAt(2).getString());
        }
    }

    @Test
    public void testReadXMLWithPhoneticHintsReadOnlySST() throws Exception {
        try (InputStream is = getResourceStream("sharedStrings-with-phonetic-hints.xml")) {
            ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(is, false);
            assertEquals(3, sst.getUniqueCount());
            assertEquals(3, sst.getCount());
            assertEquals("Country", sst.getItemAt(0).getString());
            assertEquals("City", sst.getItemAt(1).getString());
            assertEquals("沖縄", sst.getItemAt(2).getString());
        }
    }

    @Test
    public void testReadStyledXML() throws Exception {
        testReadStyledXML(false);
    }

    @Test
    public void testReadStyledXMLFullFormat() throws Exception {
        testReadStyledXML(true);
    }

    @Test
    public void testReadOOXMLStrict() throws Exception {
        testReadOOXMLStrict(false);
    }

    @Test
    public void testReadOOXMLStrictFullFormat() throws Exception {
        testReadOOXMLStrict(true);
    }

    @Test(expected = NoSuchElementException.class)
    public void testReadMissingEntry() throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true)) {
            RichTextString rts = sst.getItemAt(0);
        }
    }

    @Test(expected = NoSuchElementException.class)
    public void testGetStringMissingEntry() throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true)) {
            String str = sst.getString(0);
        }
    }

    @Test(expected = NoSuchElementException.class)
    public void testReadMissingEntryFullFormat() throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(false, true)) {
            RichTextString rts = sst.getItemAt(0);
        }
    }

    @Test(expected = NoSuchElementException.class)
    public void testGetStringMissingEntryFullFormat() throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(false, true)) {
            String str = sst.getString(0);
        }
    }

    @Test
    public void testParseMalformedCountFile() throws Exception {
        try (
                InputStream is = getResourceStream("MalformedSSTCount.xlsx");
                OPCPackage pkg = OPCPackage.open(is);
                TempFileSharedStringsTable sst = new TempFileSharedStringsTable(false, false)
        ) {
            List<PackagePart> parts = pkg.getPartsByName(Pattern.compile("/xl/sharedStrings.xml"));
            assertEquals(1, parts.size());

            SharedStringsTable stbl = new SharedStringsTable(parts.get(0));
            try (InputStream ssStream = parts.get(0).getInputStream()) {
                sst.readFrom(ssStream);
            }
            assertEquals(8, sst.getCount());
            assertEquals(stbl.getUniqueCount(), sst.getUniqueCount());
            for (int i = 0; i < stbl.getUniqueCount(); i++) {
                RichTextString i1 = stbl.getItemAt(i);
                assertEquals(i1.getString(), sst.getItemAt(i).getString());
            }
        }
    }

    @Test
    public void testWrite() throws Exception {
        testWrite(10, false);
    }

    @Test
    public void testWriteFullFormat() throws Exception {
        testWrite(10, true);
    }

    static final String MINIMAL_XML = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"55\" uniqueCount=\"49\">" +
            "<si>" +
            "<t>bla</t>" +
            "<phoneticPr fontId=\"1\"/>" +
            "</si>" +
            "</sst>";

    @Test
    public void testMinimalTable() throws IOException {
        try (TempFileSharedStringsTable tbl = new TempFileSharedStringsTable()) {
            tbl.readFrom(new ByteArrayInputStream(MINIMAL_XML.getBytes(StandardCharsets.UTF_8)));
            assertEquals(49, tbl.getUniqueCount());
            assertEquals(55, tbl.getCount());
            assertEquals("bla", tbl.getItemAt(0).getString());
            assertThrows(NoSuchElementException.class,
                    () -> tbl.getItemAt(1).getString());
        }
    }

    @Test
    public void testBigUniqueCount() throws IOException, SAXException {
        try (TempFileSharedStringsTable tbl = new TempFileSharedStringsTable()) {
            tbl.readFrom(new ByteArrayInputStream(
                    MINIMAL_XML.replace("49", Integer.toString(Integer.MAX_VALUE))
                            .getBytes(StandardCharsets.UTF_8)));
            assertNotNull(tbl);
            assertEquals(Integer.MAX_VALUE, tbl.getUniqueCount());
            assertEquals(55, tbl.getCount());
            assertEquals("bla", tbl.getItemAt(0).getString());
            assertThrows(NoSuchElementException.class,
                    () -> tbl.getItemAt(1).getString());
        }
    }

    @Test
    public void testHugeUniqueCount() throws IOException, SAXException {
        try (TempFileSharedStringsTable tbl = new TempFileSharedStringsTable()) {
            tbl.readFrom(new ByteArrayInputStream(
                    MINIMAL_XML.replace("49", "99999999999999999")
                            .getBytes(StandardCharsets.UTF_8)));
            assertNotNull(tbl);
            assertEquals(1, tbl.getUniqueCount());
            assertEquals(55, tbl.getCount());
            assertEquals("bla", tbl.getItemAt(0).getString());
            assertThrows(NoSuchElementException.class,
                    () -> tbl.getItemAt(1).getString());
        }
    }

    @Test
    public void stressTest() throws Exception {
        final int limit = 100;
        File tempFile = TempFile.createTempFile("shared-string-stress", ".tmp");
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(false, true)) {
            for (int i = 0; i < limit; i++) {
                sst.addSharedStringItem(new XSSFRichTextString(UUID.randomUUID().toString()));
            }
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                sst.writeTo(fos);
            }
            try (TempFileSharedStringsTable sst2 = new TempFileSharedStringsTable(true)) {
                try (FileInputStream fis = new FileInputStream(tempFile)){
                    sst2.readFrom(fis);
                }
                assertEquals(limit, sst2.getUniqueCount());
                assertEquals(limit, sst2.getCount());
            }
        } finally {
            tempFile.delete();
        }
    }

    private void testWrite(int size, boolean fullFormat) throws Exception {
        java.util.Random rnd = new java.util.Random();
        byte[] bytes = new byte[1028];
        try (
                UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get();
                TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)
        ) {
            for (int i = 0; i < size; i++) {
                rnd.nextBytes(bytes);
                String rndString = java.util.Base64.getEncoder().encodeToString(bytes);
                sst.addSharedStringItem(new XSSFRichTextString(rndString));
            }
            sst.writeTo(bos);
            String out = bos.toString(StandardCharsets.UTF_8);
            assertFalse("sst output should not contain xml-fragment", out.contains("xml-fragment"));
            try(TempFileSharedStringsTable sst2 = new TempFileSharedStringsTable(true, fullFormat)) {
                sst2.readFrom(bos.toInputStream());
                assertEquals(size, sst2.getCount());
            }
        }
    }

    private void testReadOOXMLStrict(boolean fullFormat) throws Exception {
        try (InputStream is = getResourceStream("strictSharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)) {
            sst.readFrom(is);
            assertEquals(15, sst.getUniqueCount());
            assertEquals(19, sst.getCount());
            assertEquals("Lorem", sst.getItemAt(0).getString());
            assertEquals("Lorem", sst.getString(0));
            assertEquals("The quick brown fox jumps over the lazy dog",
                    sst.getItemAt(14).getString());
            assertEquals("The quick brown fox jumps over the lazy dog",
                    sst.getString(14));
            int expectedFormattingRuns = fullFormat ? 11: 0;
            assertEquals(expectedFormattingRuns, sst.getItemAt(14).numFormattingRuns());
        }
    }

    private void testReadStyledXML(boolean fullFormat) throws Exception {
        try (InputStream is = getResourceStream("styledSharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)) {
            sst.readFrom(is);
            assertEquals(1, sst.getCount());
            assertEquals(1, sst.getUniqueCount());
            assertEquals("shared styled string", sst.getItemAt(0).getString());
            assertEquals("shared styled string", sst.getString(0));
        }
    }

    private void testReadXML(boolean encrypt, boolean fullFormat) throws Exception {
        try (InputStream is = getResourceStream("sharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(encrypt, fullFormat)) {
            sst.readFrom(is);
            assertEquals(60, sst.getCount());
            assertEquals(38, sst.getUniqueCount());
            assertEquals("City", sst.getItemAt(0).getString());
            assertEquals("City", sst.getString(0));
        }
    }

    private void testWriteOut(boolean fullFormat) throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)) {
            sst.addSharedStringItem(new XSSFRichTextString("First string"));
            sst.addSharedStringItem(new XSSFRichTextString("First string"));
            sst.addSharedStringItem(new XSSFRichTextString("First string"));
            sst.addSharedStringItem(new XSSFRichTextString("Second string"));
            sst.addSharedStringItem(new XSSFRichTextString("Second string"));
            sst.addSharedStringItem(new XSSFRichTextString("Second string"));
            XSSFRichTextString rts = new XSSFRichTextString("Second string");
            XSSFFont font = new XSSFFont();
            font.setFontName("Arial");
            font.setBold(true);
            rts.applyFont(font);
            sst.addSharedStringItem(rts);
            int expectedUniqueCount = fullFormat ? 3 : 2;
            assertEquals(expectedUniqueCount, sst.getUniqueCount());
            assertEquals(7, sst.getCount());
            try (UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()) {
                sst.writeTo(bos);
                try (TempFileSharedStringsTable sst2 = new TempFileSharedStringsTable(true)) {
                    sst2.readFrom(bos.toInputStream());
                    assertEquals(expectedUniqueCount, sst2.getUniqueCount());
                    assertEquals(7, sst2.getCount());
                    assertEquals("First string", sst2.getItemAt(0).getString());
                    assertEquals("First string", sst2.getString(0));
                    assertEquals("Second string", sst2.getItemAt(1).getString());
                    assertEquals("Second string", sst2.getString(1));
                    if (fullFormat) {
                        assertEquals("Second string", sst2.getItemAt(2).getString());
                        assertEquals("Second string", sst2.getString(2));
                    }
                }
                try (SharedStringsTable sst3 = new SharedStringsTable()) {
                    sst3.readFrom(bos.toInputStream());
                    assertEquals(expectedUniqueCount, sst3.getUniqueCount());
                    assertEquals(7, sst3.getCount());
                    assertEquals("First string", sst3.getItemAt(0).getString());
                    assertEquals("Second string", sst3.getItemAt(1).getString());
                    if (fullFormat) {
                        assertEquals("Second string", sst3.getItemAt(2).getString());
                    }
                }
            }
        }
    }
}
