package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.POIXMLException;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

import java.io.Serializable;

class SerializableComment implements Serializable {

    private static final long serialVersionUID = 7829136421241571165L;

    private String author;
    private String commentText;
    private boolean fullFormat = false;
    private transient CTRst ctRst; //CTRstImpl is Serializable but very inefficient
    private String addressAsText; //Serializable version of cellAddress
    private transient CellAddress cellAddress; //CellAddress is not Serializable
    private boolean visible = true;

    public SerializableComment() {}

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

    /**
     * @return comment as a rich string
     * @throws POIXMLException if the value is not parseable
     */
    public XSSFRichTextString getString() {
        return new XSSFRichTextString(getCTRst());
    }

    public String getCommentText() {
        if (fullFormat) {
            return getString().getString();
        } else {
            return commentText;
        }
    }

    public void setString(RichTextString string) {
        if(!(string instanceof XSSFRichTextString)){
            throw new IllegalArgumentException("Only XSSFRichTextString argument is supported");
        }
        ctRst = ((XSSFRichTextString)string).getCTRst();
        fullFormat = true;
        commentText = ctRst.xmlText();
    }

    public void setString(String text) {
        commentText = text;
        fullFormat = false;
    }

    private CTRst getCTRst() throws POIXMLException {
        if (ctRst == null && commentText != null) {
            //ctRst is transient so might need to be recreated from commentText
            synchronized (this) {
                if (ctRst == null && commentText != null) {
                    if (fullFormat) {
                        try {
                            ctRst = CTRst.Factory.parse(commentText);
                        } catch (XmlException e) {
                            throw new POIXMLException("Could not parse comment rich text string", e);
                        }
                    } else {
                        XSSFRichTextString richTextString = new XSSFRichTextString();
                        richTextString.setString(commentText);
                        ctRst = richTextString.getCTRst();
                    }
                }
            }
        }
        return ctRst;
    }
}
