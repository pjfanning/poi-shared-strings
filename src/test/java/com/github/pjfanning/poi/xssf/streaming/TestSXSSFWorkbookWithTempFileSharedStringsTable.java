package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

public class TestSXSSFWorkbookWithTempFileSharedStringsTable {

    @Test
    public void useStreamingSharedStringsTable() throws Exception {
        SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(new SXSSFFactory(true)),
                SXSSFWorkbook.DEFAULT_WINDOW_SIZE, true, true);

        SharedStringsTable sss = POITestCase.getFieldValue(SXSSFWorkbook.class, wb, SharedStringsTable.class, "_sharedStringSource");

        assertNotNull(sss);
        assertEquals(TempFileSharedStringsTable.class, sss.getClass());

        Row row = wb.createSheet("S1").createRow(0);

        row.createCell(0).setCellValue("A");
        row.createCell(1).setCellValue("B");
        row.createCell(2).setCellValue("A");

        XSSFWorkbook xssfWorkbook = writeOutAndReadBack(wb);
        sss = POITestCase.getFieldValue(SXSSFWorkbook.class, wb, SharedStringsTable.class, "_sharedStringSource");
        assertEquals(2, sss.getUniqueCount());
        assertTrue(wb.dispose());

        Sheet sheet1 = xssfWorkbook.getSheetAt(0);
        assertEquals("S1", sheet1.getSheetName());
        assertEquals(1, sheet1.getPhysicalNumberOfRows());
        row = sheet1.getRow(0);
        assertNotNull(row);
        Cell cell = row.getCell(0);
        assertNotNull(cell);
        assertEquals("A", cell.getStringCellValue());
        cell = row.getCell(1);
        assertNotNull(cell);
        assertEquals("B", cell.getStringCellValue());
        cell = row.getCell(2);
        assertNotNull(cell);
        assertEquals("A", cell.getStringCellValue());

        xssfWorkbook.close();
        wb.close();
    }

    XSSFWorkbook writeOutAndReadBack(Workbook wb) {
        // wb is usually an SXSSFWorkbook, but must also work on an XSSFWorkbook
        // since workbooks must be able to be written out and read back
        // several times in succession
        if(!(wb instanceof SXSSFWorkbook || wb instanceof XSSFWorkbook)) {
            throw new IllegalArgumentException("Expected an instance of SXSSFWorkbook");
        }

        XSSFWorkbook result;
        try {
            ByteArrayOutputStream baos = new ByteArrayOutputStream(8192);
            wb.write(baos);
            InputStream is = new ByteArrayInputStream(baos.toByteArray());
            result = new XSSFWorkbook(is);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        return result;
    }
}
