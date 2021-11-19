package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.xssf.streaming.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;

public class TestSXSSFWorkbookWithTempCommentsTable {
    @Test
    public void testComments() throws Exception {
        SXSSFFactory factory = new SXSSFFactory();
        factory.enableTempFileComments(true);
        SXSSFWorkbook wb = new SXSSFWorkbook(new XSSFWorkbook(factory),
                SXSSFWorkbook.DEFAULT_WINDOW_SIZE, false, false);
        try {
            SXSSFSheet sheet = wb.createSheet("testSheet");
            ClientAnchor anchor = wb.getCreationHelper().createClientAnchor();
            SXSSFRow row = sheet.createRow(0);
            SXSSFCell cell = row.createCell(0);
            cell.setCellValue("cell1");
            SXSSFDrawing drawing = sheet.createDrawingPatriarch();
            Comment comment = drawing.createCellComment(anchor);
            comment.setString(wb.getCreationHelper().createRichTextString("comment1"));
            cell.setCellComment(comment);
            try (FileOutputStream fos = new FileOutputStream("test.xlsx")) {
                wb.write(fos);
            }
        } finally {
            wb.close();
            wb.dispose();
        }
    }
}
