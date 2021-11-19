package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import java.io.Serializable;

class SerializableComment implements Serializable {

    private static final long serialVersionUID = 7829136421241571165L;

    private String author;
    private CTRst ctRst; //CTRstImpl is Serializable
    private String addressAsText; //Serializable version of cellAddress
    private transient CellAddress cellAddress; //CellAddress is not Serializable
    private boolean visible = true;

    public SerializableComment() {

    }

    public void setAddress(CellAddress address) {
        this.cellAddress = address;
        this.addressAsText = address.formatAsString();
    }

    public String getAuthor() {
        return author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public int getColumn() {
        CellAddress address = getAddress();
        if (address == null) {
            throw new IllegalStateException("cell address not initialised");
        }
        return address.getColumn();
    }

    public int getRow() {
        CellAddress address = getAddress();
        if (address == null) {
            throw new IllegalStateException("cell address not initialised");
        }
        return address.getRow();
    }

    public boolean isVisible() {
        return visible;
    }

    public void setVisible(boolean visible) {
        this.visible = visible;
    }

    public CellAddress getAddress() {
        if (cellAddress == null && addressAsText != null) {
            //cellAddress is transient so might need to be recreated from addressAsText
            cellAddress = new CellAddress(addressAsText);
        }
        return cellAddress;
    }

    public void setAddress(int row, int col) {
        setAddress(new CellAddress(row, col));
    }

    public void setColumn(int col) {
        CellAddress address = getAddress();
        if (address == null) {
            throw new IllegalStateException("cell address not initialised");
        }
        setAddress(address.getRow(), col);
    }

    public void setRow(int row) {
        CellAddress address = getAddress();
        if (address == null) {
            throw new IllegalStateException("cell address not initialised");
        }
        setAddress(row, address.getColumn());
    }

    public XSSFRichTextString getString() {
        return new XSSFRichTextString(ctRst);
    }

    public void setString(RichTextString string) {
        if(!(string instanceof XSSFRichTextString)){
            throw new IllegalArgumentException("Only XSSFRichTextString argument is supported");
        }
        this.ctRst = ((XSSFRichTextString)string).getCTRst();
    }

    public void setString(String text) {
        XSSFRichTextString rts = new XSSFRichTextString();
        rts.setString(text);
        setString(rts);
    }
}
