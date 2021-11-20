package com.github.pjfanning.poi.xssf.streaming;

import org.apache.commons.io.output.UnsynchronizedByteArrayOutputStream;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.xssf.streaming.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

import java.io.FileOutputStream;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;

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
            try (UnsynchronizedByteArrayOutputStream bos = new UnsynchronizedByteArrayOutputStream()) {
                wb.write(bos);
                try (XSSFWorkbook wb2 = new XSSFWorkbook(bos.toInputStream())) {
                    XSSFSheet xssfSheet = wb2.getSheetAt(0);
                    XSSFRow xssfRow = xssfSheet.getRow(0);
                    XSSFCell xssfCell = xssfRow.getCell(0);
                    assertEquals(cell.getStringCellValue(), xssfCell.getStringCellValue());
                    Comment xssfComment = cell.getCellComment();
                    assertNotNull("xssfComment found?", xssfComment);
                    assertEquals(comment.getString().getString(), xssfComment.getString().getString());
                }
            }
        } finally {
            wb.close();
            wb.dispose();
        }
    }
}
