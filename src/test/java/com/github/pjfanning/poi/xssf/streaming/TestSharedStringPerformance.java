package com.github.pjfanning.poi.xssf.streaming;

import com.github.pjfanning.poi.xssf.streaming.cache.CachedSharedStringsTable;
import com.github.pjfanning.poi.xssf.streaming.cache.lru.SSTCacheLRU;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.Test;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

public class TestSharedStringPerformance {
    private static Logger log = LoggerFactory.getLogger(TestSharedStringPerformance.class);
    private static int SIZE = 100000;

    @Test
    public void testH2() throws Exception {
        try (TempFileSharedStringsTable sst = new TempFileSharedStringsTable(false, false)) {
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

    @Test
    public void testLRU() throws Exception {
        try (
                SSTCacheLRU cache = new SSTCacheLRU(100);
                CachedSharedStringsTable sst = new CachedSharedStringsTable(cache, true)
        ) {
            log.info("lru starting write phase");
            long writeStartTime = System.currentTimeMillis();
            for(int i = 0; i < SIZE; i++) {
                sst.addSharedStringItem(new XSSFRichTextString(Integer.toString(i)));
            }
            log.info("lru finished write phase time={}ms", (System.currentTimeMillis() - writeStartTime));
            log.info("lru starting read phase");
            long readStartTime = System.currentTimeMillis();
            for(int i = 0; i < SIZE; i++) {
                RichTextString rts = sst.getItemAt(i);
                if (rts == null) {
                    log.info("lru unexpected null pos={}", i);
                }
            }
            log.info("lru finished read phase time={}ms", (System.currentTimeMillis() - readStartTime));
        }
    }

    @Test
    public void testPOISST() throws Exception {
        try (SharedStringsTable sst = new SharedStringsTable()) {
            log.info("poi starting write phase");
            long writeStartTime = System.currentTimeMillis();
            for(int i = 0; i < SIZE; i++) {
                sst.addSharedStringItem(new XSSFRichTextString(Integer.toString(i)));
            }
            log.info("poi finished write phase time={}ms", (System.currentTimeMillis() - writeStartTime));
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
