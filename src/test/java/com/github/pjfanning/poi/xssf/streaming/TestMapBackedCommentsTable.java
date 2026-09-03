package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.collections4.IteratorUtils;
import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Assert;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.UUID;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import static com.github.pjfanning.poi.xssf.streaming.TestIOUtils.getResourceStream;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;
import static org.junit.Assert.assertTrue;

public class TestMapBackedCommentsTable {

    @Test
    public void testAuthorSetViaDelegateIsDeclaredInOutput() throws Exception {
        // findAuthor registers unseen authors, so resolving them while writing the comments
        // used to append an author after <authors> had already been written, leaving the
        // comment pointing at an authorId that was never declared
        try (MapBackedCommentsTable ct = new MapBackedCommentsTable(false)) {
            ct.createNewComment(new XSSFClientAnchor(0, 0, 0, 0, 1, 1, 2, 2))
                    .setString(new XSSFRichTextString("hello"));
            // DelegatingXSSFComment.setAuthor bypasses the POI code that registers the author
            ct.findCellComment(new CellAddress(1, 1)).setAuthor("Alice");

            final String xml;
            try (UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()) {
                ct.writeTo(bos);
                xml = new String(bos.toByteArray(), StandardCharsets.UTF_8);
            }

            final int declaredAuthors = xml.split("<author>", -1).length - 1;
            assertTrue("expected at least one <author> to be declared", declaredAuthors > 0);
            assertTrue("expected Alice to be declared as an author", xml.contains("<author>Alice</author>"));

            final Matcher matcher = Pattern.compile("authorId=\"(\\d+)\"").matcher(xml);
            assertTrue("expected a comment with an authorId", matcher.find());
            do {
                final int authorId = Integer.parseInt(matcher.group(1));
                assertTrue("authorId " + authorId + " is out of range, only " + declaredAuthors
                        + " author(s) declared", authorId < declaredAuthors);
            } while (matcher.find());
        }
    }

    @Test
    public void testSpecialCharactersInAuthorsRoundTrip() throws Exception {
        // the write path escapes authors itself rather than going through XMLBeans,
        // so check the awkward characters survive a write/read round trip
        final String[] authors = {
            "plain",
            "amp & author",
            "angle < and >",
            "quote \" and apos '",
            "cdata ]]> close",
            "unicode \u00e9\u65e5"
        };
        final String[] texts = {
            "plain text",
            "text with & ampersand",
            "text with < and > angles",
            "text with \"quotes\" and 'apostrophes'",
            "text with ]]> in it",
            "text with \u00e9\u65e5 unicode"
        };

        final String xml;
        try (MapBackedCommentsTable ct = new MapBackedCommentsTable(false)) {
            for (int i = 0; i < authors.length; i++) {
                final XSSFComment comment = ct.createNewComment(
                        new XSSFClientAnchor(0, 0, 0, 0, 0, i, 1, i + 1));
                comment.setString(new XSSFRichTextString(texts[i]));
                comment.setAuthor(authors[i]);
            }
            try (UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get()) {
                ct.writeTo(bos);
                xml = new String(bos.toByteArray(), StandardCharsets.UTF_8);
            }
        }

        try (MapBackedCommentsTable readBack = new MapBackedCommentsTable(false)) {
            readBack.readFrom(new ByteArrayInputStream(xml.getBytes(StandardCharsets.UTF_8)));
            assertEquals("number of comments", authors.length, readBack.getNumberOfComments());
            for (int i = 0; i < authors.length; i++) {
                final XSSFComment comment = readBack.findCellComment(new CellAddress(i, 0));
                assertNotNull("comment at row " + i, comment);
                assertEquals("author at row " + i, authors[i], comment.getAuthor());
                assertEquals("text at row " + i, texts[i], comment.getString().getString());
            }
        }
    }

    @Test
    public void testReadXML() throws Exception {
        testReadXML(false);
    }

    @Test
    public void testReadXMLWithFullFormat() throws Exception {
        testReadXML(true);
    }

    @Test
    public void testWriteEmpty() throws Exception {
        try (
                UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get();
                MapBackedCommentsTable commentsTable = new MapBackedCommentsTable(true)
        ) {
            commentsTable.writeTo(bos);
            try (MapBackedCommentsTable commentsTable2 = new MapBackedCommentsTable(false)) {
                commentsTable2.readFrom(bos.toInputStream());
                assertEquals(0, commentsTable2.getNumberOfComments());
            }
        }
    }

    @Test
    public void testWrite() throws Exception {
        testWrite(false);
    }

    @Test
    public void testWriteWithFullFormat() throws Exception {
        testWrite(true);
    }

    @Test
    public void testReadStrictXML() throws Exception {
        try (
                InputStream is = getResourceStream("strict-comments1.xml");
                MapBackedCommentsTable ct = new MapBackedCommentsTable(false)
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

            XSSFComment testComment = ct.findCellComment(addresses.get(0));
            XSSFRichTextString testCommentText = testComment.getString();
            assertTrue("comment text contains expected value?",
                    testCommentText.getString().contains("Gaeilge"));
            assertEquals(0, testCommentText.numFormattingRuns());
            //not set because there is no sheet data and related VMLDrawing data
            assertNull("client anchor not set?", testComment.getClientAnchor());
        }
    }

    @Test
    public void testMoveComment() throws Exception {
        try (
                SXSSFWorkbook workbook = new SXSSFWorkbook();
                MapBackedCommentsTable commentsTable = new MapBackedCommentsTable(true)
        ) {
            CreationHelper factory = workbook.getCreationHelper();
            SXSSFSheet sheet = workbook.createSheet();
            commentsTable.setSheet(sheet);
            SXSSFRow row = sheet.createRow(0);
            SXSSFCell cell = row.createCell(0);
            ClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(0);
            anchor.setCol2(1);
            anchor.setRow1(row.getRowNum());
            anchor.setRow2(row.getRowNum());
            XSSFComment comment = commentsTable.createNewComment(anchor);
            String uniqueText = UUID.randomUUID().toString();
            comment.setString(uniqueText);
            comment.setAuthor("author" + uniqueText);

            XSSFComment comment1 = commentsTable.findCellComment(new CellAddress("A1"));
            assertEquals(comment.getString().getString(), comment1.getString().getString());

            comment.setAddress(1, 1);
            assertNull("no longer a comment on cell A1?", commentsTable.findCellComment(new CellAddress("A1")));

            XSSFComment comment2 = commentsTable.findCellComment(new CellAddress("B2"));
            assertEquals(comment.getString().getString(), comment2.getString().getString());
        }
    }

    @Test
    public void testMoveCommentCopy() throws Exception {
        try (
                SXSSFWorkbook workbook = new SXSSFWorkbook();
                MapBackedCommentsTable commentsTable = new MapBackedCommentsTable(true)
        ) {
            CreationHelper factory = workbook.getCreationHelper();
            SXSSFSheet sheet = workbook.createSheet();
            commentsTable.setSheet(sheet);
            SXSSFRow row = sheet.createRow(0);
            SXSSFCell cell = row.createCell(0);
            ClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(0);
            anchor.setCol2(1);
            anchor.setRow1(row.getRowNum());
            anchor.setRow2(row.getRowNum());
            XSSFComment comment = commentsTable.createNewComment(anchor);
            String uniqueText = UUID.randomUUID().toString();
            comment.setString(uniqueText);
            comment.setAuthor("author" + uniqueText);

            XSSFComment comment1 = commentsTable.findCellComment(new CellAddress("A1"));
            assertEquals(comment.getString().getString(), comment1.getString().getString());

            //like testMoveComment but moves the copy of the comment (comment1) instead
            comment1.setAddress(1, 1);
            assertNull("no longer a comment on cell A1?", commentsTable.findCellComment(new CellAddress("A1")));

            XSSFComment comment2 = commentsTable.findCellComment(new CellAddress("B2"));
            assertEquals(comment.getString().getString(), comment2.getString().getString());
        }
    }

    @Test
    public void testModifyComment() throws Exception {
        try (
                SXSSFWorkbook workbook = new SXSSFWorkbook();
                MapBackedCommentsTable commentsTable = new MapBackedCommentsTable(true)
        ) {
            CreationHelper factory = workbook.getCreationHelper();
            SXSSFSheet sheet = workbook.createSheet();
            commentsTable.setSheet(sheet);
            SXSSFRow row = sheet.createRow(0);
            SXSSFCell cell = row.createCell(0);
            ClientAnchor anchor = factory.createClientAnchor();
            anchor.setCol1(0);
            anchor.setCol2(1);
            anchor.setRow1(row.getRowNum());
            anchor.setRow2(row.getRowNum());
            XSSFComment comment = commentsTable.createNewComment(anchor);
            comment.setString("initText");
            comment.setAuthor("initAuthor");

            XSSFComment comment1 = commentsTable.findCellComment(new CellAddress("A1"));
            assertEquals(comment.getString().getString(), comment1.getString().getString());
            assertEquals(comment.getAuthor(), comment1.getAuthor());

            String uniqueText = UUID.randomUUID().toString();
            comment.setString(uniqueText);
            comment.setAuthor("author" + uniqueText);

            XSSFComment comment2 = commentsTable.findCellComment(new CellAddress("A1"));
            assertEquals(comment.getString().getString(), comment2.getString().getString());
            assertEquals(comment.getAuthor(), comment2.getAuthor());
        }
    }

    @Test
    public void testReadXMLWithPhoneticHints() throws Exception {
        try (InputStream is = getResourceStream("sharedStrings-with-phonetic-hints.xml");
             MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(false)) {
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
    public void stressTest() throws Exception {
        final int limit = 100;
        File tempFile = TempFile.createTempFile("comments-stress", ".tmp");
        try (
                SXSSFWorkbook workbook = new SXSSFWorkbook();
                MapBackedCommentsTable commentsTable = new MapBackedCommentsTable(true)
        ) {
            CreationHelper factory = workbook.getCreationHelper();
            SXSSFSheet sheet = workbook.createSheet();
            commentsTable.setSheet(sheet);
            for (int i = 0; i < limit; i++) {
                SXSSFRow row = sheet.createRow(i);
                SXSSFCell cell = row.createCell(0);
                ClientAnchor anchor = factory.createClientAnchor();
                anchor.setCol1(0);
                anchor.setCol2(1);
                anchor.setRow1(row.getRowNum());
                anchor.setRow2(row.getRowNum());
                XSSFComment comment = commentsTable.createNewComment(anchor);
                String uniqueText = UUID.randomUUID().toString();
                comment.setString(uniqueText);
                comment.setAuthor("author" + uniqueText);
            }
            try (FileOutputStream fos = new FileOutputStream(tempFile)) {
                commentsTable.writeTo(fos);
            }
            try (
                    FileInputStream fis = new FileInputStream(tempFile);
                    MapBackedCommentsTable commentsTable2 = new MapBackedCommentsTable(true)
            ) {
                commentsTable2.readFrom(fis);
                //also includes the empty author that is automatically added
                assertEquals(limit + 1, commentsTable2.getNumberOfAuthors());
                assertEquals(0, commentsTable2.findAuthor(""));
                assertEquals(limit, commentsTable2.getNumberOfComments());
            }
        } finally {
            tempFile.delete();
        }
    }

    private void testReadXML(boolean fullFormat) throws Exception {
        try (
                InputStream is = getResourceStream("comments1.xml");
                MapBackedCommentsTable ct = new MapBackedCommentsTable(fullFormat)
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

            DelegatingXSSFComment delegatingXSSFComment = (DelegatingXSSFComment)ct.findCellComment(addresses.get(0));
            assertNull("ctShape not set?", delegatingXSSFComment.getCTShape());
            XSSFRichTextString testComment = delegatingXSSFComment.getString();
            assertEquals("comment top row1 (index0)",
                    testComment.getString()
                            .replace("\n", "").replace("\r", ""));
            int expectedRuns = fullFormat ? 2 : 0;
            assertEquals(expectedRuns, testComment.numFormattingRuns());
            assertEquals("comment top row1 (index0)",
                delegatingXSSFComment.getCommentText().replace("\n", "").replace("\r", ""));
        }
    }

    private void testWrite(boolean fullFormat) throws Exception {
        try (
                UnsynchronizedByteArrayOutputStream bos = UnsynchronizedByteArrayOutputStream.builder().get();
                InputStream is = getResourceStream("comments1.xml");
                MapBackedCommentsTable commentsTable = new MapBackedCommentsTable(fullFormat)
        ) {
            commentsTable.readFrom(is);
            assertEquals(3, commentsTable.getNumberOfComments());
            commentsTable.writeTo(bos);
            String out = bos.toString(StandardCharsets.UTF_8);
            assertFalse("XML must not contain xml-fragment element", out.contains("xml-fragment"));
            try (MapBackedCommentsTable commentsTable2 = new MapBackedCommentsTable(fullFormat)) {
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

}
