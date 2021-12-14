package com.github.pjfanning.poi.xssf.streaming.cache.lru;

import com.github.pjfanning.poi.xssf.streaming.cache.SSTCache;
import org.apache.poi.util.NotImplemented;
import org.apache.poi.util.TempFile;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import java.io.File;
import java.util.Iterator;
import java.util.function.Consumer;

public class SSTCacheLRU implements SSTCache, AutoCloseable {

    private final FileBackedList strings;
    private File tempFile;

    public SSTCacheLRU(int capacity) {
        try {
            tempFile = TempFile.createTempFile("poi-shared-strings", ".tmp");
            strings = new FileBackedList(tempFile, capacity);
        } catch (Exception e) {
            if (tempFile != null) {
                tempFile.delete();
            }
            throw new RuntimeException(e);
        }
    }

    @Override
    public CTRst putCTRst(Integer idx, CTRst st) {
        strings.add(st.xmlText());
        return st;
    }

    @Override
    public CTRst getCTRst(Integer idx) {
        return new XSSFRichTextString(strings.getAt(idx)).getCTRst();
    }

    @Override
    public Integer putStringIndex(String s, Integer idx) {
        // Doing nothing, because this implementation doesn't support storing (String, Integer) key value pairs.
        return idx;
    }

    @Override
    @NotImplemented
    public Integer getStringIndex(String s) {
        throw new IllegalStateException("The SSTCacheLRU implementation doesn't allow checking for String keys.");
    }

    @Override
    public boolean containsString(String s) {
        // Always returning false, because this implementation doesn't support storing (String, Integer)
        // key value pairs. It results in every entry to be considered as unique.
        return false;
    }

    @Override
    public Iterator<Integer> keyIterator() {
        return new LRUCacheKeyIterator(strings);
    }

    @Override
    public void close() {
        tempFile.delete();
    }

    public static class Builder {
        private int capacity = 1000;

        public Builder cacheSizeBytes(int capacity) {
            this.capacity = capacity;
            return this;
        }

        public SSTCacheLRU build() {
            return new SSTCacheLRU(capacity);
        }
    }

    private static class LRUCacheKeyIterator implements Iterator<Integer> {

        private final FileBackedList fileBackedList;
        private Integer counter = -1;

        public LRUCacheKeyIterator(FileBackedList fileBackedList) {
            this.fileBackedList = fileBackedList;
        }

        @Override
        public boolean hasNext() {
            return counter < fileBackedList.size() - 1;
        }

        @Override
        public Integer next() {
            counter++;
            return counter;
        }

        @Override
        public void remove() {
            throw new IllegalStateException("LRUCacheKeyIterator doesn't support this operation.");
        }

        @Override
        public void forEachRemaining(Consumer action) {
            throw new IllegalStateException("LRUCacheKeyIterator doesn't support this operation.");
        }
    }
}
