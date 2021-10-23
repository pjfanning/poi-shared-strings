package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.io.PrintWriter;
import java.io.StringWriter;

public class BasicSheetContentsHandler implements XSSFSheetXMLHandler.SheetContentsHandler {

    StringWriter sw = new StringWriter();
    PrintWriter pw = new PrintWriter(sw);

    @Override
    public void startRow(int rowNum) {
        pw.println("Start Row " + rowNum);
    }

    @Override
    public void endRow(int rowNum) {
        pw.println("End Row " + rowNum);
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        pw.print(cellReference);
        pw.print(' ');
        pw.print(formattedValue);
        if (comment != null && comment.getString() != null) {
            pw.print(" Comment=");
            pw.print(comment.getString());
        }
        pw.println();
    }

    void println(String value) {
        pw.println(value);
    }

    String getExtract() {
        pw.close();
        return sw.toString();
    }
}
