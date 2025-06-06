package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Test;
import org.xml.sax.SAXException;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.NoSuchElementException;
import java.util.UUID;

import static com.github.pjfanning.poi.xssf.streaming.TestIOUtils.getResourceStream;
import static com.github.pjfanning.poi.xssf.streaming.TestTempFileSharedStringsTable.MINIMAL_XML;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertThrows;

public class TestMapBackedSharedStringsTable {
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
        testReadXML(false);
    }

    @Test
    public void testReadXMLFullFormat() throws Exception {
        testReadXML(true);
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
        try (MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable()) {
            RichTextString rts = sst.getItemAt(0);
        }
    }

    @Test(expected = NoSuchElementException.class)
    public void testGetStringMissingEntry() throws Exception {
        try (MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable()) {
            String str = sst.getString(0);
        }
    }

    @Test(expected = NoSuchElementException.class)
    public void testReadMissingEntryFullFormat() throws Exception {
        try (MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(true)) {
            RichTextString rts = sst.getItemAt(0);
        }
    }

    @Test(expected = NoSuchElementException.class)
    public void testGetStringMissingEntryFullFormat() throws Exception {
        try (MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(true)) {
            String str = sst.getString(0);
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

    @Test
    public void testMinimalTable() throws IOException {
        try (MapBackedSharedStringsTable tbl = new MapBackedSharedStringsTable()) {
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
        try (MapBackedSharedStringsTable tbl = new MapBackedSharedStringsTable()) {
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
        try (MapBackedSharedStringsTable tbl = new MapBackedSharedStringsTable()) {
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
        try (MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(true)) {
            for (int i = 0; i < limit; i++) {
                sst.addSharedStringItem(new XSSFRichTextString(UUID.randomUUID().toString()));
            }
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                sst.writeTo(fos);
            }
            try (MapBackedSharedStringsTable sst2 = new MapBackedSharedStringsTable(true)) {
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
                MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(fullFormat)
        ) {
            for (int i = 0; i < size; i++) {
                rnd.nextBytes(bytes);
                String rndString = java.util.Base64.getEncoder().encodeToString(bytes);
                sst.addSharedStringItem(new XSSFRichTextString(rndString));
            }
            sst.writeTo(bos);
            String out = bos.toString(StandardCharsets.UTF_8);
            assertFalse("sst output should not contain xml-fragment", out.contains("xml-fragment"));
            try(MapBackedSharedStringsTable sst2 = new MapBackedSharedStringsTable(fullFormat)) {
                sst2.readFrom(bos.toInputStream());
                assertEquals(size, sst2.getCount());
            }
        }
    }

    private void testReadOOXMLStrict(boolean fullFormat) throws Exception {
        try (InputStream is = getResourceStream("strictSharedStrings.xml");
             MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(fullFormat)) {
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
             MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(fullFormat)) {
            sst.readFrom(is);
            assertEquals(1, sst.getCount());
            assertEquals(1, sst.getUniqueCount());
            assertEquals("shared styled string", sst.getItemAt(0).getString());
            assertEquals("shared styled string", sst.getString(0));
        }
    }

    private void testReadXML(boolean fullFormat) throws Exception {
        try (InputStream is = getResourceStream("sharedStrings.xml");
             MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(fullFormat)) {
            sst.readFrom(is);
            assertEquals(60, sst.getCount());
            assertEquals(38, sst.getUniqueCount());
            assertEquals("City", sst.getItemAt(0).getString());
            assertEquals("City", sst.getString(0));
        }
    }

    private void testWriteOut(boolean fullFormat) throws Exception {
        try (MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(fullFormat)) {
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
                try (MapBackedSharedStringsTable sst2 = new MapBackedSharedStringsTable(true)) {
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
