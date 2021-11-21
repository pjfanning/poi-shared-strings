package com.github.pjfanning.poi.xssf.streaming;

import com.microsoft.schemas.vml.CTShape;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.NotImplemented;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTComment;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRst;

public class DelegatingXSSFComment extends XSSFComment {
    private final SerializableComment delegate;
    private final Comments comments;
    private final CTShape ctShape;

    DelegatingXSSFComment(Comments comments, SerializableComment delegate, CTShape ctShape) {
        super(comments, null, null);
        this.comments = comments;
        this.delegate = delegate;
        this.ctShape = ctShape;
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

    @Override
    public CTComment getCTComment() {
        CTRst rst = delegate.getString().getCTRst();
        CTComment ctComment = CTComment.Factory.newInstance();
        ctComment.setText(rst);
        return ctComment;
    }

    @Override
    public void setAddress(int row, int col) {
        CellAddress oldAddress = delegate.getAddress();
        delegate.setAddress(row, col);
        comments.referenceUpdated(oldAddress, this);
    }

    @Override
    public void setAddress(CellAddress address) {
        CellAddress oldAddress = delegate.getAddress();
        delegate.setAddress(address);
        comments.referenceUpdated(oldAddress, this);
    }

    @Override
    public void setRow(int row) {
        CellAddress oldAddress = delegate.getAddress();
        delegate.setRow(row);
        comments.referenceUpdated(oldAddress, this);
    }

    @Override
    public void setColumn(int col) {
        CellAddress oldAddress = delegate.getAddress();
        delegate.setColumn(col);
        comments.referenceUpdated(oldAddress, this);
    }

    @Override
    public void setString(RichTextString string) {
        delegate.setString(string);
        comments.commentUpdated(this);
    }

    @Override
    public void setString(String string) {
        delegate.setString(string);
        comments.commentUpdated(this);
    }

    @Override
    public void setAuthor(String author) {
        delegate.setAuthor(author);
        comments.commentUpdated(this);
    }

    @Override
    public void setVisible(boolean visible) {
        delegate.setVisible(visible);
        comments.commentUpdated(this);
    }

    @Override
    protected CTShape getCTShape() {
        return ctShape;
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
    
}
