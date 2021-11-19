package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Assert;
import org.junit.Test;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.UUID;

import static org.junit.Assert.*;

public class TestTempFileCommentsTable {

    @Test
    public void testReadXML() throws Exception {
        testReadXML(false, false);
    }

    @Test
    public void testReadXMLWithEncryptedTempFile() throws Exception {
        testReadXML(true, false);
    }

    @Test
    public void testReadXMLWithFullFormat() throws Exception {
        testReadXML(false, true);
    }

    @Test
    public void testWriteEmpty() throws Exception {
        try (
                UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream();
                TempFileCommentsTable commentsTable = new TempFileCommentsTable(true)
        ) {
            commentsTable.writeTo(bos);
            try (TempFileCommentsTable commentsTable2 = new TempFileCommentsTable(false)) {
                commentsTable2.readFrom(bos.toInputStream());
                assertEquals(0, commentsTable2.getNumberOfComments());
            }
        }
    }

    @Test
    public void testWrite() throws Exception {
        testWrite(false, false);
    }

    @Test
    public void testWriteWithEncryptedTempFile() throws Exception {
        testWrite(true, false);
    }

    @Test
    public void testWriteWithFullFormat() throws Exception {
        testWrite(false, true);
    }

    @Test
    public void testReadStrictXML() throws Exception {
        try (
                InputStream is = TestTempFileCommentsTable.class.getClassLoader().getResourceAsStream("strict-comments1.xml");
                TempFileCommentsTable ct = new TempFileCommentsTable(false, false)
        ) {
            ct.readFrom(is);
            assertEquals(1, ct.getNumberOfComments());
            List<CellAddress> addresses = IteratorUtils.toList(ct.getCellAddresses());
            assertEquals(1, addresses.size());
            assertEquals("B1", addresses.get(0).formatAsString());
            for (CellAddress address : addresses) {
                Assert.assertNotNull(ct.findCellComment(address));
            }

            assertEquals(1, ct.getNumberOfAuthors());
            assertEquals("tc={12222C35-D781-4D4A-81D9-2C6FD97BD160}", ct.getAuthor(0));
            assertEquals(1, ct.findAuthor("new-author"));
            assertEquals("new-author", ct.getAuthor(1));

            XSSFRichTextString testComment = ct.findCellComment(addresses.get(0)).getString();
            assertTrue("comment text contains expected value?",
                    testComment.getString().contains("Gaeilge"));
            assertEquals(0, testComment.numFormattingRuns());
        }
    }

    private void testReadXML(boolean encrypt, boolean fullFormat) throws Exception {
        try (
                InputStream is = TestTempFileCommentsTable.class.getClassLoader().getResourceAsStream("comments1.xml");
                TempFileCommentsTable ct = new TempFileCommentsTable(encrypt, fullFormat)
        ) {
            ct.readFrom(is);
            assertEquals(3, ct.getNumberOfComments());
            List<CellAddress> addresses = IteratorUtils.toList(ct.getCellAddresses());
            assertEquals(3, addresses.size());
            assertEquals("A1", addresses.get(0).formatAsString());
            assertEquals("A3", addresses.get(1).formatAsString());
            assertEquals("A4", addresses.get(2).formatAsString());
            for (CellAddress address : addresses) {
                Assert.assertNotNull(ct.findCellComment(address));
            }

            assertEquals(1, ct.getNumberOfAuthors());
            assertEquals("Sven Nissel", ct.getAuthor(0));
            assertEquals(1, ct.findAuthor("new-author"));
            assertEquals("new-author", ct.getAuthor(1));

            XSSFRichTextString testComment = ct.findCellComment(addresses.get(0)).getString();
            assertEquals("comment top row1 (index0)",
                    testComment.getString()
                            .replace("\n", "").replace("\r", ""));
            int expectedRuns = fullFormat ? 2 : 0;
            assertEquals(expectedRuns, testComment.numFormattingRuns());
        }
    }

    private void testWrite(boolean encrypt, boolean fullFormat) throws Exception {
        try (
                UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream();
                InputStream is = TestTempFileCommentsTable.class.getClassLoader().getResourceAsStream("comments1.xml");
                TempFileCommentsTable commentsTable = new TempFileCommentsTable(encrypt, fullFormat)
        ) {
            commentsTable.readFrom(is);
            assertEquals(3, commentsTable.getNumberOfComments());
            commentsTable.writeTo(bos);
            String out = bos.toString(StandardCharsets.UTF_8);
            assertFalse("XML must not contain xml-fragment element", out.contains("xml-fragment"));
            try (TempFileCommentsTable commentsTable2 = new TempFileCommentsTable(false, fullFormat)) {
                commentsTable2.readFrom(bos.toInputStream());
                assertEquals(3, commentsTable2.getNumberOfComments());
                List<CellAddress> addresses = IteratorUtils.toList(commentsTable2.getCellAddresses());
                assertEquals(3, addresses.size());
                assertEquals("A1", addresses.get(0).formatAsString());
                assertEquals("A3", addresses.get(1).formatAsString());
                assertEquals("A4", addresses.get(2).formatAsString());
                for (CellAddress address : addresses) {
                    Assert.assertNotNull(commentsTable2.findCellComment(address));
                }

                assertEquals(1, commentsTable2.getNumberOfAuthors());
                assertEquals("Sven Nissel", commentsTable2.getAuthor(0));
                assertEquals(1, commentsTable2.findAuthor("new-author"));
                assertEquals("new-author", commentsTable2.getAuthor(1));
            }
            CommentsTable commentsTable3 = new CommentsTable();
            commentsTable3.readFrom(bos.toInputStream());
            assertEquals(3, commentsTable3.getNumberOfComments());
            List<CellAddress> addresses = IteratorUtils.toList(commentsTable3.getCellAddresses());
            assertEquals(3, addresses.size());
            assertEquals("A1", addresses.get(0).formatAsString());
            assertEquals("A3", addresses.get(1).formatAsString());
            assertEquals("A4", addresses.get(2).formatAsString());
            for (CellAddress address : addresses) {
                Assert.assertNotNull(commentsTable3.findCellComment(address));
            }

            assertEquals(1, commentsTable3.getNumberOfAuthors());
            assertEquals("Sven Nissel", commentsTable3.getAuthor(0));
            assertEquals(1, commentsTable3.findAuthor("new-author"));
            assertEquals("new-author", commentsTable3.getAuthor(1));
        }
    }

    @Test
    public void stressTest() throws Exception {
        final int limit = 10000;
        try (
                SXSSFWorkbook workbook = new SXSSFWorkbook();
                TempFileCommentsTable commentsTable = new TempFileCommentsTable(false, true)
        ) {
            CreationHelper factory = workbook.getCreationHelper();
            SXSSFSheet sheet = workbook.createSheet();
            for (int i = 0; i < limit; i++) {
                SXSSFRow row = sheet.createRow(i);
                SXSSFCell cell = row.createCell(0);
                ClientAnchor anchor = factory.createClientAnchor();
                anchor.setCol1(0);
                anchor.setCol2(1);
                anchor.setRow1(row.getRowNum());
                anchor.setRow2(row.getRowNum());
                XSSFComment comment = commentsTable.createNewComment(sheet, anchor);
                commentsTable.getAuthor()
                comment.setString(UUID.randomUUID().toString());
                comment.setAuthor();
            }
        }
    }

}
