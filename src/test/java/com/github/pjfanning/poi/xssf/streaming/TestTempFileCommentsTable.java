package com.github.pjfanning.poi.xssf.streaming;

import org.junit.Assert;
import org.junit.Test;

import java.io.InputStream;

public class TestTempFileCommentsTable {

    @Test
    public void testReadXML() throws Exception {
        try (InputStream is = TestTempFileCommentsTable.class.getClassLoader().getResourceAsStream("comments1.xml");
             TempFileCommentsTable ct = new TempFileCommentsTable(true)) {
            ct.readFrom(is);
            Assert.assertEquals(3, ct.getNumberOfComments());
        }
    }
}
