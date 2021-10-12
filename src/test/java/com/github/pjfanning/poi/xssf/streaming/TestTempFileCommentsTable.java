package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.util.CellAddress;
import org.junit.Assert;
import org.junit.Test;

import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;

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
        }
    }

}
