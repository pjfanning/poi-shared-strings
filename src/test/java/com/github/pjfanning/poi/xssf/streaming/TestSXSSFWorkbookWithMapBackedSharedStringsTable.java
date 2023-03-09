package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

public class TestSXSSFWorkbookWithMapBackedSharedStringsTable {

    @Test
    public void useStreamingSharedStringsTable() throws Exception {
        final SharedStringsTableFactory sharedStringsTableFactory0 =
                () -> new MapBackedSharedStringsTable(false);
        final SharedStringsTableFactory sharedStringsTableFactory1 =
                () -> new MapBackedSharedStringsTable(true);
        XSSFFactory factory0 = CustomXSSFFactory.builder()
                .sharedStringsFactory(sharedStringsTableFactory0)
                .build();
        XSSFFactory factory1 = CustomXSSFFactory.builder()
                .sharedStringsFactory(sharedStringsTableFactory1)
                .build();
        for (XSSFFactory factory : new XSSFFactory[]{factory0, factory1}) {
            try (SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(factory),
                        SXSSFWorkbook.DEFAULT_WINDOW_SIZE, true, true)) {
                SharedStringsTable sss = POITestUtils.getFieldValue(SXSSFWorkbook.class, wb, SharedStringsTable.class, "_sharedStringSource");

                assertNotNull(sss);

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
            }
        }
    }
}
