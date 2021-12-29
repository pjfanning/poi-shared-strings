package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

public class TestSXSSFWorkbookWithTempFileSharedStringsTable {

    @Test
    public void useStreamingSharedStringsTable() throws Exception {
        SXSSFFactory factory0 = new SXSSFFactory();
        SXSSFFactory factory1 = new SXSSFFactory().encryptTempFiles(true);
        for (SXSSFFactory factory : new SXSSFFactory[]{factory0, factory1}) {
            SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(factory),
                    SXSSFWorkbook.DEFAULT_WINDOW_SIZE, true, true);

            SharedStringsTable sss = POITestUtils.getFieldValue(SXSSFWorkbook.class, wb, SharedStringsTable.class, "_sharedStringSource");

            assertNotNull(sss);
            assertEquals(TempFileSharedStringsTable.class, sss.getClass());

            Row row = wb.createSheet("S1").createRow(0);

            row.createCell(0).setCellValue("A");
            row.createCell(1).setCellValue("B");
            row.createCell(2).setCellValue("A");

            XSSFWorkbook xssfWorkbook = POITestUtils.writeOutAndReadBack(wb);
            sss = POITestUtils.getFieldValue(SXSSFWorkbook.class, wb, SharedStringsTable.class, "_sharedStringSource");
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
    }
}
