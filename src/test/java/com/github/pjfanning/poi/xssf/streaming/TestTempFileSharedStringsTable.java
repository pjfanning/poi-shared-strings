package com.github.pjfanning.poi.xssf.streaming;

import java.io.*;

import org.apache.poi.ooxml.util.DocumentHelper;
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
                DocumentHelper.newDocumentBuilder().parse(new ByteArrayInputStream(bos.toByteArray()));
            }
        }
    }

    @Test
    public void testReadXML() throws Exception {
        try (InputStream is = TestTempFileSharedStringsTable.class.getClassLoader().getResourceAsStream("sharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true)) {
            sst.readFrom(is);
            Assert.assertEquals(38, sst.getCount());
            Assert.assertEquals("City", sst.getItemAt(0).getString());
        }
    }


    @Test
    public void testReadStyledXML() throws Exception {
        try (InputStream is = TestTempFileSharedStringsTable.class.getClassLoader().getResourceAsStream("styledSharedStrings.xml");
             TempFileSharedStringsTable sst = new TempFileSharedStringsTable(true)) {
            sst.readFrom(is);
            Assert.assertEquals(1, sst.getCount());
            Assert.assertEquals("shared styled string", sst.getItemAt(0).getString());
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
}
