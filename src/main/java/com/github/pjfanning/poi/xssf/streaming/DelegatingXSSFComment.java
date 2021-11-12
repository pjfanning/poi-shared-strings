package com.github.pjfanning.poi.xssf.streaming;

import com.microsoft.schemas.vml.CTShape;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;

public class DelegatingXSSFComment extends XSSFComment {
    private final SerializableComment delegate;

    public DelegatingXSSFComment(SerializableComment delegate) {
        super(null, null, null);
        this.delegate = delegate;
    }

    @Override
    public String getAuthor() {
        return delegate.getAuthor();
    }

    @Override
    public void setAuthor(String author) {
        delegate.setAuthor(author);
    }

    @Override
    public int getColumn() {
        return delegate.getColumn();
    }

    @Override
    public int getRow() {
        return delegate.getRow();
    }

    @Override
    public boolean isVisible() {
        return delegate.isVisible();
    }

    @Override
    public void setVisible(boolean visible) {
        delegate.setVisible(visible);
    }

    @Override
    public CellAddress getAddress() {
        return delegate.getAddress();
    }

    @Override
    public void setAddress(int row, int col) {
        delegate.setAddress(row, col);
    }

    @Override
    public void setAddress(CellAddress address) {
        delegate.setAddress(address);
    }

    @Override
    public void setColumn(int col) {
        delegate.setColumn(col);
    }

    @Override
    public void setRow(int row) {
        delegate.setRow(row);
    }

    @Override
    public XSSFRichTextString getString() {
        return delegate.getString();
    }

    @Override
    public void setString(RichTextString string) {
        delegate.setString(string);
    }

    @Override
    public void setString(String string) {
        delegate.setString(string);
    }

    @Override
    public ClientAnchor getClientAnchor() {
        throw new RuntimeException("Not Implemented");
    }

    @Override
    protected CTComment getCTComment() {
        throw new RuntimeException("Not Implemented");
    }

    @Override
    protected CTShape getCTShape() {
        throw new RuntimeException("Not Implemented");
    }
}
