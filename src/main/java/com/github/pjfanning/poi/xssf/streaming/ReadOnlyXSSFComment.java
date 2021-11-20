package com.github.pjfanning.poi.xssf.streaming;

import com.microsoft.schemas.vml.CTShape;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.NotImplemented;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;

public class ReadOnlyXSSFComment extends XSSFComment {
    private final SerializableComment delegate;

    public ReadOnlyXSSFComment(SerializableComment delegate) {
        super(null, null, null);
        this.delegate = delegate;
    }

    @Override
    public String getAuthor() {
        return delegate.getAuthor();
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
    public CellAddress getAddress() {
        return delegate.getAddress();
    }

    @Override
    public XSSFRichTextString getString() {
        return delegate.getString();
    }

    /**
     * Not implemented. This class only supports read-only methods.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public void setAddress(int row, int col) {
        throw new IllegalStateException("Not Implemented");
    }

    /**
     * Not implemented. This class only supports read-only methods.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public void setAddress(CellAddress address) {
        throw new IllegalStateException("Not Implemented");
    }

    /**
     * Not implemented. This class only supports read-only methods.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public void setRow(int row) {
        throw new IllegalStateException("update actions are not supported");
    }

    /**
     * Not implemented. This class only supports read-only methods.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public void setColumn(int col) {
        throw new IllegalStateException("Not Implemented");
    }

    /**
     * Not implemented. This class only supports read-only methods.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public void setString(RichTextString string) {
        throw new IllegalStateException("update actions are not supported");
    }

    /**
     * Not implemented. This class only supports read-only methods.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public void setString(String string) {
        throw new IllegalStateException("update actions are not supported");
    }

    /**
     * Not implemented. This class only supports read-only methods.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public void setAuthor(String author) {
        throw new IllegalStateException("update actions are not supported");
    }

    /**
     * Not implemented.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public void setVisible(boolean visible) {
        throw new IllegalStateException("Not Implemented");
    }

    /**
     * Not implemented.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    public ClientAnchor getClientAnchor() {
        throw new IllegalStateException("Not Implemented");
    }

    /**
     * Not implemented.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    protected CTComment getCTComment() {
        throw new IllegalStateException("Not Implemented");
    }

    /**
     * Not implemented.
     * @throws IllegalStateException
     */
    @Override
    @NotImplemented
    protected CTShape getCTShape() {
        throw new IllegalStateException("Not Implemented");
    }
}
