package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.poi.ooxml.util.DocumentHelper;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Assert;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.List;

public class TestTempFileCommentsTable {

    @Test
    public void testReadXML() throws Exception {
        try (InputStream is = TestTempFileCommentsTable.class.getClassLoader().getResourceAsStream("comments1.xml");
             TempFileCommentsTable ct = new TempFileCommentsTable(true)) {
            ct.readFrom(is);
            Assert.assertEquals(3, ct.getNumberOfComments());
            List<CellAddress> addresses = IteratorUtils.toList(ct.getCellAddresses());
            Assert.assertEquals(3, addresses.size());
            Assert.assertEquals("A1", addresses.get(0).formatAsString());
            Assert.assertEquals("A3", addresses.get(1).formatAsString());
            Assert.assertEquals("A4", addresses.get(2).formatAsString());
            for (CellAddress address : addresses) {
                Assert.assertNotNull(ct.findCellComment(address));
            }

            Assert.assertEquals(1, ct.getNumberOfAuthors());
            Assert.assertEquals("Sven Nissel", ct.getAuthor(0));
            Assert.assertEquals(1, ct.findAuthor("new-author"));
            Assert.assertEquals("new-author", ct.getAuthor(1));
        }
    }

    @Test
    public void testWriteOut() throws Exception {
        try (TempFileCommentsTable commentsTable = new TempFileCommentsTable(true)) {

        }
    }
}
