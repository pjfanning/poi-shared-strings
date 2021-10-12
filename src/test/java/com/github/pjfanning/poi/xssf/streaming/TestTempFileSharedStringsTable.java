package com.github.pjfanning.poi.xssf.streaming;

import java.io.*;
import java.util.NoSuchElementException;

import org.apache.poi.ooxml.util.DocumentHelper;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Assert;
import org.junit.Test;

public class TestTempFileSharedStringsTable {
    @Test
    public void testWriteOut() throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true)) {
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
            Assert.assertEquals(3, sst.getUniqueCount());
            Assert.assertEquals(7, sst.getCount());
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                sst.writeTo(bos);
                try (TempFileSharedStringsTable sst2 = new TempFileSharedStringsTable(true)) {
                    sst2.readFrom(new ByteArrayInputStream(bos.toByteArray()));
                    Assert.assertEquals(3, sst.getUniqueCount());
                    Assert.assertEquals(7, sst.getCount());
                    Assert.assertEquals("First string", sst.getItemAt(0).getString());
                    Assert.assertEquals("Second string", sst.getItemAt(1).getString());
                    Assert.assertEquals("Second string", sst.getItemAt(2).getString());
                }
            }
        }
    }

    @Test
    public void testReadXML() throws Exception {
        try (InputStream is = TestTempFileSharedStringsTable.class.getClassLoader().getResourceAsStream("sharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true)) {
            sst.readFrom(is);
            Assert.assertEquals(60, sst.getCount());
            Assert.assertEquals(38, sst.getUniqueCount());
            Assert.assertEquals("City", sst.getItemAt(0).getString());
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

    @Test
    public void testWrite() throws Exception {
        testWrite(10);
    }

    private void testWrite(int size) throws Exception {
        java.util.Random rnd = new java.util.Random();
        byte[] bytes = new byte[1028];
        File file = new File("sst.txt");
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true)) {
            for (int i = 0; i < size; i++) {
                rnd.nextBytes(bytes);
                String rndString = java.util.Base64.getEncoder().encodeToString(bytes);
                sst.addSharedStringItem(new XSSFRichTextString(rndString));
            }
            try (java.io.FileOutputStream fos = new FileOutputStream(file)) {
                sst.writeTo(fos);
            }
        } finally {
            file.delete();
        }
    }

    private void testReadOOXMLStrict(boolean fullFormat) throws Exception {
        try (InputStream is = TestTempFileSharedStringsTable.class.getClassLoader().getResourceAsStream("strictSharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)) {
            sst.readFrom(is);
            Assert.assertEquals(15, sst.getUniqueCount());
            Assert.assertEquals(19, sst.getCount());
            Assert.assertEquals("Lorem", sst.getItemAt(0).getString());
            Assert.assertEquals("The quick brown fox jumps over the lazy dog",
                    sst.getItemAt(14).getString());
        }
    }

    private void testReadStyledXML(boolean fullFormat) throws Exception {
        try (InputStream is = TestTempFileSharedStringsTable.class.getClassLoader().getResourceAsStream("styledSharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true, fullFormat)) {
            sst.readFrom(is);
            Assert.assertEquals(1, sst.getCount());
            Assert.assertEquals(1, sst.getUniqueCount());
            Assert.assertEquals("shared styled string", sst.getItemAt(0).getString());
        }
    }
}
