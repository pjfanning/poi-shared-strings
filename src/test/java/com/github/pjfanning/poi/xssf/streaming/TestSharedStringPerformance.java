package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TestSharedStringPerformance {
    private static Logger log = LoggerFactory.getLogger(TestSharedStringPerformance.class);
    private static int SIZE = 100000;

    @Test
    public void testH2() throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(false, true)) {
            log.info("h2 starting write phase");
            long writeStartTime = System.currentTimeMillis();
            for(int i = 0; i < SIZE; i++) {
                sst.addSharedStringItem(new XSSFRichTextString(Integer.toString(i)));
            }
            log.info("h2 finished write phase time={}ms", (System.currentTimeMillis() - writeStartTime));
            log.info("h2 starting read phase");
            long readStartTime = System.currentTimeMillis();
            for(int i = 0; i < SIZE; i++) {
                RichTextString rts = sst.getItemAt(i);
                if (rts == null) {
                    log.info("h2 unexpected null pos={}", i);
                }
            }
            log.info("h2 finished read phase time={}ms", (System.currentTimeMillis() - readStartTime));
        }
    }

    @Test
    public void testMapDB() throws Exception {
        try (MapDBSharedStringsTable sst = new MapDBSharedStringsTable(true)) {
            log.info("mapdb starting write phase");
            long writeStartTime = System.currentTimeMillis();
            for(int i = 0; i < SIZE; i++) {
                sst.addSharedStringItem(new XSSFRichTextString(Integer.toString(i)));
            }
            log.info("mapdb finished write phase time={}ms", (System.currentTimeMillis() - writeStartTime));
            log.info("mapdb starting read phase");
            long readStartTime = System.currentTimeMillis();
            for(int i = 0; i < SIZE; i++) {
                RichTextString rts = sst.getItemAt(i);
                if (rts == null) {
                    log.info("mapdb unexpected null pos={}", i);
                }
            }
            log.info("mapdb finished read phase time={}ms", (System.currentTimeMillis() - readStartTime));
        }
    }

}
