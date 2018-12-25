package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;

class SimpleXSSFComment extends XSSFComment {

    private String author;
    private String text;
    private CellAddress address;
    private boolean visible = true;

    public SimpleXSSFComment() {
        super(null, null, null);
    }

    @Override
    public void setAddress(CellAddress address) {
        this.address = address;
    }

    @Override
    public String getAuthor() {
        return author;
    }

    @Override
    public void setAuthor(String author) {
        this.author = author;
    }

    @Override
    public int getColumn() {
        if (address == null) {
            throw new IllegalStateException("cell address not initialised");
        }
        return address.getColumn();
    }

    @Override
    public int getRow() {
        if (address == null) {
            throw new IllegalStateException("cell address not initialised");
        }
        return address.getRow();
    }

    @Override
    public boolean isVisible() {
        return visible;
    }

    @Override
    public void setVisible(boolean visible) {
        this.visible = visible;
    }

    @Override
    public CellAddress getAddress() {
        return address;
    }

    @Override
    public void setAddress(int row, int col) {
        setAddress(new CellAddress(row, col));
    }

    @Override
    public void setColumn(int col) {
        if (address == null) {
            throw new IllegalStateException("cell address not initialised");
        }
        setAddress(address.getRow(), col);
    }

    @Override
    public void setRow(int row) {
        if (address == null) {
            throw new IllegalStateException("cell address not initialised");
        }
        setAddress(row, address.getColumn());
    }

    @Override
    public XSSFRichTextString getString() {
        return new XSSFRichTextString(text);
    }

    @Override
    public void setString(RichTextString string) {
        setString(string.getString());
    }

    @Override
    public void setString(String text) {
        this.text = text;
    }
}
