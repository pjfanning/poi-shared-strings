package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.util.PackageHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.junit.Assert;
import org.junit.Test;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintWriter;
import java.io.StringWriter;

public class TestStreamingRead {

    @Test
    public void testXSSFSheetXMLHandler() throws Exception {
        try (InputStream is = TestIOUtils.getResourceStream("sample.xlsx");
             OPCPackage pkg = PackageHelper.open(is);
             TempFileSharedStringsTable strings = new TempFileSharedStringsTable(pkg, true)) {
            XSSFReader xssfReader = new XSSFReader(pkg);
            StylesTable styles = xssfReader.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
            int index = 0;
            BasicSheetContentsHandler handler = new BasicSheetContentsHandler();
            while (iter.hasNext()) {
                try (InputStream stream = iter.next()) {
                    String sheetName = iter.getSheetName();
                    handler.println(sheetName + " [index=" + index + "]:");
                    processSheet(styles, strings, handler, stream);
                }
                ++index;
            }
            Assert.assertEquals(getExpected(), handler.getExtract());
        }
    }

    private void processSheet(
            StylesTable styles,
            SharedStrings strings,
            SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try (TempFileCommentsTable commentsTable = new TempFileCommentsTable(true)) {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, commentsTable, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch(ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    private String getExpected() {
        StringWriter sw = new StringWriter();
        try (PrintWriter pw = new PrintWriter(sw)) {
            pw.println("Sheet1 [index=0]:");
            pw.println("Start Row 0");
            pw.println("A1 12345");
            pw.println("End Row 0");
            pw.println("Start Row 1");
            pw.println("A2 abcdef");
            pw.println("End Row 1");
            pw.println("Start Row 2");
            pw.println("A3 \uD835\uDF4A\uD835\uDF4B\uD835\uDF4C\uD835\uDF4D\uD835\uDF4E\uD835\uDF4F\uD835\uDF50\uD835\uDF51\uD835\uDF52\uD835\uDF53\uD835\uDF54\uD835\uDF55\uD835\uDF56\uD835\uDF57\uD835\uDF58\uD835\uDF59\uD835\uDF5A\uD835\uDF5B\uD835\uDF5C\uD835\uDF5D\uD835\uDF5E\uD835\uDF5F\uD835\uDF60\uD835\uDF61\uD835\uDF62\uD835\uDF63\uD835\uDF64\uD835\uDF65\uD835\uDF66\uD835\uDF67\uD835\uDF68\uD835\uDF69\uD835\uDF6A\uD835\uDF6B\uD835\uDF6C\uD835\uDF6D\uD835\uDF6E\uD835\uDF6F\uD835\uDF70\uD835\uDF71\uD835\uDF72\uD835\uDF73\uD835\uDF74\uD835\uDF75\uD835\uDF76\uD835\uDF77\uD835\uDF78\uD835\uDF79\uD835\uDF7A");
            pw.println("End Row 2");
        }
        return sw.toString();
    }
}
