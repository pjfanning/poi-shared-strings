package com.github.pjfanning.poi.xssf.streaming;

import java.io.InputStream;

final class TestIOUtils {
    static InputStream getResourceStream(String filename) {
        return TestIOUtils.class.getClassLoader().getResourceAsStream(filename);
    }
}
