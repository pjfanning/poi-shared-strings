package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.util.PackageHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.junit.Test;

import java.io.InputStream;

public class TestStreamingRead {

    @Test
    public void testXSSFSheetXMLHandler() throws Exception {
        try (InputStream is = getResourceStream("sample.xlsx");
             OPCPackage pkg = PackageHelper.open(is);
             TempFileSharedStringsTable strings = new TempFileSharedStringsTable(pkg, true)) {
            XSSFReader reader = new XSSFReader(pkg);
            new XSSFSheetXMLHandler(reader.getStylesTable(), strings, createSheetContentsHandler(), false);
        }
    }

    private static XSSFSheetXMLHandler.SheetContentsHandler createSheetContentsHandler() {
        return new XSSFSheetXMLHandler.SheetContentsHandler() {

            @Override
            public void startRow(int rowNum) {
            }

            @Override
            public void endRow(int rowNum) {
            }

            @Override
            public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            }
        };
    }

    private InputStream getResourceStream(String filename) {
        return TestStreamingRead.class.getClassLoader().getResourceAsStream(filename);
    }
}
