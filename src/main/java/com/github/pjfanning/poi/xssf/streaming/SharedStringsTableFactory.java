package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.xssf.model.SharedStrings;

public interface SharedStringsTableFactory {

    /**
     * @return a new {@link SharedStrings} implementation instance, configured to your requirements
     */
    SharedStrings createSharedStringsTable();
}
