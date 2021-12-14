package com.github.pjfanning.poi.xssf.streaming.cache;

import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import java.util.Iterator;

public interface SSTCache {

    CTRst putCTRst(Integer idx, CTRst st);
    CTRst getCTRst(Integer idx);
    Integer putStringIndex(String s, Integer idx);
    Integer getStringIndex(String s);
    boolean containsString(String s);
    Iterator<Integer> keyIterator();
    void close();

}
