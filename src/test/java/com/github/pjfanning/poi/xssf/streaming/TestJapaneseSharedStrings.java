package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.junit.Test;

import java.io.InputStream;

import static org.junit.Assert.assertEquals;

public class TestJapaneseSharedStrings {
    @Test
    public void testReadOnlySharedStringsTable() throws Exception {
        try (InputStream stream = TestIOUtils.getResourceStream("jaSharedStrings.xml")) {
            ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(stream);
            assertEquals("売上 ウリアゲ ", sst.getItemAt(0).getString());
        }
    }

    @Test
    public void testReadOnlySharedStringsTableIgnorePhoneticRuns() throws Exception {
        try (InputStream stream = TestIOUtils.getResourceStream("jaSharedStrings.xml")) {
            ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(stream, false);
            assertEquals("売上", sst.getItemAt(0).getString());
        }
    }

    @Test
    public void testPoiSharedStringsTable() throws Exception {
        try (
                InputStream stream = TestIOUtils.getResourceStream("jaSharedStrings.xml");
                SharedStringsTable sst = new SharedStringsTable();
        ) {
            sst.readFrom(stream);
            assertEquals("売上", sst.getItemAt(0).getString());
        }
    }
}
