package com.github.pjfanning.poi.xssf.streaming;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.NoSuchElementException;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

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
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true)) {
            RichTextString rts = sst.getItemAt(0);
        }
    }

    @Test(expected = NoSuchElementException.class)
    public void testReadMissingEntryFullFormat() throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(false, true)) {
            RichTextString rts = sst.getItemAt(0);
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

    private void testWrite(int size, boolean fullFormat) throws Exception {
        java.util.Random rnd = new java.util.Random();
        byte[] bytes = new byte[1028];
        try (
                UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream();
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
        try (InputStream is = TestTempFileSharedStringsTable.class.getClassLoader().getResourceAsStream("strictSharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)) {
            sst.readFrom(is);
            assertEquals(15, sst.getUniqueCount());
            assertEquals(19, sst.getCount());
            assertEquals("Lorem", sst.getItemAt(0).getString());
            assertEquals("The quick brown fox jumps over the lazy dog",
                    sst.getItemAt(14).getString());
            int expectedFormattingRuns = fullFormat ? 11: 0;
            assertEquals(expectedFormattingRuns, sst.getItemAt(14).numFormattingRuns());
        }
    }

    private void testReadStyledXML(boolean fullFormat) throws Exception {
        try (InputStream is = TestTempFileSharedStringsTable.class.getClassLoader().getResourceAsStream("styledSharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)) {
            sst.readFrom(is);
            assertEquals(1, sst.getCount());
            assertEquals(1, sst.getUniqueCount());
            assertEquals("shared styled string", sst.getItemAt(0).getString());
        }
    }

    private void testReadXML(boolean fullFormat) throws Exception {
        try (InputStream is = TestTempFileSharedStringsTable.class.getClassLoader().getResourceAsStream("sharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)) {
            sst.readFrom(is);
            assertEquals(60, sst.getCount());
            assertEquals(38, sst.getUniqueCount());
            assertEquals("City", sst.getItemAt(0).getString());
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
            assertEquals(3, sst.getUniqueCount());
            assertEquals(7, sst.getCount());
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                sst.writeTo(bos);
                try (TempFileSharedStringsTable sst2 = new TempFileSharedStringsTable(true)) {
                    sst2.readFrom(new ByteArrayInputStream(bos.toByteArray()));
                    assertEquals(3, sst.getUniqueCount());
                    assertEquals(7, sst.getCount());
                    assertEquals("First string", sst.getItemAt(0).getString());
                    assertEquals("Second string", sst.getItemAt(1).getString());
                    assertEquals("Second string", sst.getItemAt(2).getString());
                }
                try (SharedStringsTable sst3 = new SharedStringsTable()) {
                    sst3.readFrom(new ByteArrayInputStream(bos.toByteArray()));
                    assertEquals(3, sst.getUniqueCount());
                    assertEquals(7, sst.getCount());
                    assertEquals("First string", sst.getItemAt(0).getString());
                    assertEquals("Second string", sst.getItemAt(1).getString());
                    assertEquals("Second string", sst.getItemAt(2).getString());
                }
            }
        }
    }

}
