package com.github.pjfanning.poi.xssf.streaming;

import fi.iki.elonen.NanoHTTPD;
import org.apache.poi.ooxml.util.PackageHelper;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Test;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.*;
import java.util.function.Consumer;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.fail;

public class TestEntityExpansion {

    @Test
    public void testEntityExpansion() {
        ExploitServer.withServer(s -> fail("Should not have made request"), () -> {
            try (InputStream is = getResourceStream("entity-expansion-exploit-poc-file.xlsx");
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
            } catch (RuntimeException re) {
                throw re;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        });
    }

    @Test
    public void testXSSFEntityExpansion() {
        ExploitServer.withServer(s -> fail("Should not have made request"), () -> {
            try (InputStream is = getResourceStream("entity-expansion-exploit-poc-file.xlsx");
                 XSSFWorkbook wb = new XSSFWorkbook(is)) {
                for (int i = 0; i < wb.getNumberOfSheets(); i++) {
                    for (Row row : wb.getSheetAt(i)) {
                        for (Cell cell : row) {
                            assertNotNull(cell.getCellType());
                        }
                    }
                }
            } catch (RuntimeException re) {
                throw re;
            } catch (Exception e) {
                throw new RuntimeException(e);
            }
        });
    }

    private void processSheet(
            StylesTable styles,
            SharedStrings strings,
            SheetContentsHandler sheetHandler,
            InputStream sheetInputStream) throws IOException, SAXException {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(
                    styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch(ParserConfigurationException e) {
            throw new RuntimeException("SAX parser appears to be broken - " + e.getMessage());
        }
    }

    private InputStream getResourceStream(String filename) {
        return TestEntityExpansion.class.getClassLoader().getResourceAsStream(filename);
    }

    private String getExpected() {
        StringWriter sw = new StringWriter();
        try (PrintWriter pw = new PrintWriter(sw)) {
            pw.println("Sheet1 [index=0]:");
            pw.println("Start Row 0");
            pw.println("A1 poc");
            pw.println("End Row 0");
        }
        return sw.toString();
    }

    private static class ExploitServer extends NanoHTTPD implements AutoCloseable {
        private final Consumer<IHTTPSession> onRequest;

        public ExploitServer(Consumer<IHTTPSession> onRequest) throws IOException {
            super(61932);
            this.onRequest = onRequest;
        }

        @Override
        public Response serve(IHTTPSession session) {
            onRequest.accept(session);
            return newFixedLengthResponse("<!ENTITY % data SYSTEM \"file://pom.xml\">\n");
        }

        public static void withServer(Consumer<IHTTPSession> onRequest, Runnable func) {
            try(ExploitServer server = new ExploitServer(onRequest)) {
                server.start(NanoHTTPD.SOCKET_READ_TIMEOUT, false);
                func.run();
            } catch(IOException e) {
                throw new UncheckedIOException(e);
            }
        }

        @Override
        public void close() {
            this.stop();
        }
    }
}
