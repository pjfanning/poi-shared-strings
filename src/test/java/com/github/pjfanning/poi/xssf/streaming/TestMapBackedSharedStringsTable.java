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
import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.NoSuchElementException;
import java.util.Set;
import java.util.UUID;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;

import static com.github.pjfanning.poi.xssf.streaming.TestIOUtils.getResourceStream;
import static com.github.pjfanning.poi.xssf.streaming.TestTempFileSharedStringsTable.MINIMAL_XML;
import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertFalse;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertThrows;
import static org.junit.Assert.assertTrue;

public class TestMapBackedSharedStringsTable {
    @Test
    public void testConcurrentAddSharedStringItem() throws Exception {
        // count/uniqueCount are plain ints and the dedupe was a check-then-act on stmap, so
        // concurrent adds could hand out the same index to two different strings
        final int threads = 8;
        final int perThread = 2000;
        final int total = threads * perThread;
        try (MapBackedSharedStringsTable sst = new MapBackedSharedStringsTable(false)) {
            final ExecutorService pool = Executors.newFixedThreadPool(threads);
            try {
                final CountDownLatch start = new CountDownLatch(1);
                final List<Future<?>> futures = new ArrayList<>();
                for (int t = 0; t < threads; t++) {
                    final int threadNo = t;
                    futures.add(pool.submit(() -> {
                        start.await();
                        for (int i = 0; i < perThread; i++) {
                            sst.addSharedStringItem(new XSSFRichTextString("t" + threadNo + "-s" + i));
                        }
                        return null;
                    }));
                }
                start.countDown();
                for (Future<?> future : futures) {
                    future.get();
                }
            } finally {
                pool.shutdown();
            }

            assertEquals("count", total, sst.getCount());
            assertEquals("uniqueCount", total, sst.getUniqueCount());
            // every index from 0 to total-1 must hold exactly one of the strings we added
            final Set<String> seen = new HashSet<>();
            for (int i = 0; i < total; i++) {
                final String value = sst.getString(i);
                assertNotNull("no entry at index " + i, value);
                assertTrue("index " + i + " holds a duplicate: " + value, seen.add(value));
            }
            assertEquals("distinct strings stored", total, seen.size());
        }
    }

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
