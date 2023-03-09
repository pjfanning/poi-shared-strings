package com.github.pjfanning.poi.xssf.streaming;

import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ooxml.POIXMLRelation;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.model.CommentsTable;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.usermodel.XSSFFactory;
import org.apache.poi.xssf.usermodel.XSSFRelation;

/**
 * Can be used with {@link org.apache.poi.xssf.usermodel.XSSFWorkbook} or {@link org.apache.poi.xssf.streaming.SXSSFWorkbook}
 * constructors to override the implementation of {@link org.apache.poi.xssf.model.Comments} and/or
 * {@link org.apache.poi.xssf.model.SharedStrings}. Use {@link #builder()} to build your factory instance.
 *
 * @see SXSSFFactory an alternative factory class that is easier to use but limited in the type of table implememtations it supports
 * @see TempFileSharedStringsTable
 * @see TempFileCommentsTable
 * @see MapBackedSharedStringsTable
 * @see MapBackedCommentsTable
 * @see SharedStringsTable
 * @see CommentsTable
 */
public class CustomXSSFFactory extends XSSFFactory {

    public static class Builder {
        private CommentsTableFactory commentsTableFactory;
        private SharedStringsTableFactory sharedStringsFactory;

        /**
         * @param commentsTableFactory a comment table instance
         * @return this Builder instance
         */
        public Builder commentsTableFactory(CommentsTableFactory commentsTableFactory) {
            this.commentsTableFactory = commentsTableFactory;
            return this;
        }

        /**
         * @param sharedStringsFactory a sharedStringsFactory table instance
         * @return this Builder instance
         */
        public Builder sharedStringsFactory(SharedStringsTableFactory sharedStringsFactory) {
            this.sharedStringsFactory = sharedStringsFactory;
            return this;
        }

        public CustomXSSFFactory build() {
            return new CustomXSSFFactory(commentsTableFactory, sharedStringsFactory);
        }
    }

    public static Builder builder() {
        return new Builder();
    }

    private final CommentsTableFactory commentsTableFactory;
    private final SharedStringsTableFactory sharedStringsFactory;


    private CustomXSSFFactory(CommentsTableFactory commentsTableFactory, SharedStringsTableFactory sharedStringsFactory) {
        this.commentsTableFactory = commentsTableFactory;
        this.sharedStringsFactory = sharedStringsFactory;
    }

    /**
     * @param descriptor  describes the object to create
     * @return a {@link POIXMLDocumentPart} that is created using the build factories (if set)
     * @throws FactoryMismatchException if the created instances do not implement {@link POIXMLDocumentPart}
     */
    @Override
    public POIXMLDocumentPart newDocumentPart(POIXMLRelation descriptor) {
        if (XSSFRelation.SHARED_STRINGS.getRelation().equals(descriptor.getRelation())
                && sharedStringsFactory != null) {
            final SharedStrings sharedStrings = sharedStringsFactory.createSharedStringsTable();
            if (sharedStrings instanceof POIXMLDocumentPart) {
                return (POIXMLDocumentPart) sharedStrings;
            } else if (sharedStrings != null) {
                throw new FactoryMismatchException("Shared Strings Table must implement POIXMLDocumentPart");
            }
            return null;
        }
        if (XSSFRelation.SHEET_COMMENTS.getRelation().equals(descriptor.getRelation())
                && commentsTableFactory != null) {
            Comments commentsTable = commentsTableFactory.createCommentsTable();
            if (commentsTable instanceof POIXMLDocumentPart) {
                return (POIXMLDocumentPart) commentsTable;
            } else if (commentsTable != null) {
                throw new FactoryMismatchException("Comments Table must implement POIXMLDocumentPart");
            }
            return null;
        }
        return super.newDocumentPart(descriptor);
    }
}
