package com.github.pjfanning.poi.xssf.streaming;

import org.junit.Test;

import static org.junit.Assert.assertEquals;

public class TestWriteUtils {
    @Test
    public void testStripXmlFragmentElement() throws Exception{
        assertEquals("<root/>", WriteUtils.stripXmlFragmentElement("<root/>"));
        assertEquals("<root/>",
                WriteUtils.stripXmlFragmentElement("<xml-fragment><root/></xml-fragment>"));
        assertEquals("<root/>",
                WriteUtils.stripXmlFragmentElement("<xml-fragment xmlns:main=\"https://main.com\"><root/></xml-fragment>"));
    }
}
